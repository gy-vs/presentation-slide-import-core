"""Deep-copy slides, and the package parts they depend on, between presentations.

This is the machinery behind :meth:`Slides.import_slides`. A slide is copied together
with a self-contained copy of every part it references (its slide-layout and that
layout's slide-master, images, charts, external hyperlinks, etc.) so the destination
package can be opened without the source package present.

Relationships inside the copied graph are rebuilt within the destination package.
Internal slide-jump actions are rewritten to point at the imported copies; a jump to a
slide that was not selected for import raises rather than silently producing a deck
that jumps to the wrong slide.
"""

from __future__ import annotations

import posixpath
import re
from typing import TYPE_CHECKING, Iterable, cast

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart

if TYPE_CHECKING:
    from pptx.package import Package
    from pptx.parts.presentation import PresentationPart


class SlideImportError(ValueError):
    """Raised when one or more slides cannot be imported as a self-contained unit.

    Subclasses |ValueError| so callers that already catch |ValueError| keep working.
    """


class _SlideImporter:
    """Copy specified slide parts (and their dependencies) into a destination package."""

    def __init__(
        self,
        dst_presentation_part: PresentationPart,
        source_slide_parts: Iterable[SlidePart],
    ):
        self._dst_pres_part = dst_presentation_part
        self._dst_package: Package = dst_presentation_part.package
        self._source_slide_parts = tuple(source_slide_parts)
        # -- source SlidePart -> copied destination SlidePart, in requested order --
        self._slide_copies: dict[SlidePart, SlidePart] = {}
        # -- every copied part: source Part -> copied Part --
        self._part_copies: dict[Part, Part] = {}
        # -- partnames handed out during this import, for collision-free allocation --
        self._allocated_partnames: set[str] = set()
        # -- reserved sequential partnames for the explicitly selected slides --
        self._reserved_slide_partnames: dict[SlidePart, PackURI] = {}

    @classmethod
    def import_slides(
        cls,
        dst_presentation_part: PresentationPart,
        source_slide_parts: Iterable[SlidePart],
    ) -> tuple[SlidePart, ...]:
        """Return tuple of new |SlidePart| objects imported into the destination package.

        The returned slides appear in the same order as `source_slide_parts`. Raises
        |SlideImportError| if the operation cannot produce a self-contained result; in that
        case the destination package is left unchanged.
        """
        return cls(dst_presentation_part, source_slide_parts)._perform()

    # =================================================================================
    # orchestration
    # =================================================================================

    def _perform(self) -> tuple[SlidePart, ...]:
        self._validate_source_slides()

        # -- assign slide partnames up-front, continuing the destination's slide
        # -- numbering so imported slides get tidy names like slide3.xml, slide4.xml --
        existing_slide_count = self._dst_pres_part.slide_count
        for idx, source_slide_part in enumerate(self._source_slide_parts):
            partname = PackURI("/ppt/slides/slide%d.xml" % (existing_slide_count + idx + 1))
            self._allocated_partnames.add(str(partname))
            self._reserved_slide_partnames[source_slide_part] = partname

        # -- Phase 1: construct (but do not yet relate) a copy of every required part.
        # -- Validation happens here; until phase 2 the destination package is untouched.
        for source_slide_part in self._source_slide_parts:
            self._copy_slide_part(source_slide_part)

        # -- Phase 2: reconstruct relationships inside the destination package --
        for source_part, copied_part in self._part_copies.items():
            self._copy_relationships(source_part, copied_part)

        # -- Phase 3: register copied masters and the new slides with the presentation --
        self._register_slide_masters()
        slide_parts = self._register_slides()
        return slide_parts

    def _validate_source_slides(self) -> None:
        """Raise if the requested import cannot be carried out."""
        source_slide_parts = self._source_slide_parts

        if len(source_slide_parts) == 0:
            raise SlideImportError("no slides were supplied to import")

        for source_slide_part in source_slide_parts:
            if not isinstance(source_slide_part, SlidePart):
                raise SlideImportError(
                    "each source slide must be a Slide object from the source presentation"
                )

        source_packages = {part.package for part in source_slide_parts}
        if len(source_packages) > 1:
            raise SlideImportError(
                "all imported slides must belong to the same source presentation"
            )

        source_package = next(iter(source_packages))
        if source_package is self._dst_package:
            raise SlideImportError(
                "cannot import slides from the presentation they are being imported into"
            )

        if len(set(source_slide_parts)) != len(source_slide_parts):
            raise SlideImportError("same source slide specified more than once in one import")

    # =================================================================================
    # phase 1 -- copy parts
    # =================================================================================

    def _copy_slide_part(self, source_slide_part: SlidePart) -> SlidePart:
        """Return a partless copy of `source_slide_part`, copying its dependencies first."""
        existing = self._part_copies.get(source_slide_part)
        if existing is not None:
            return cast(SlidePart, existing)

        if source_slide_part not in self._selected_source_slides:
            # -- reached through a slide-to-slide relationship that was not selected --
            raise SlideImportError(
                "slide has a click action that targets a slide which was not "
                "selected for import; import that slide too or remove the action"
            )

        copied_slide_part = cast(SlidePart, self._construct_copy(source_slide_part))
        self._part_copies[source_slide_part] = copied_slide_part
        self._slide_copies[source_slide_part] = copied_slide_part

        for rel in source_slide_part.rels.values():
            if rel.is_external or rel.reltype == RT.NOTES_SLIDE:
                continue
            self._copy_part_reached_from_slide(rel.reltype, rel.target_part)

        return copied_slide_part

    def _copy_part_reached_from_slide(self, reltype: str, source_part: Part) -> None:
        """Copy `source_part` reached by a relationship of type `reltype` from a slide."""
        if reltype == RT.SLIDE:
            self._copy_slide_part(cast(SlidePart, source_part))
        elif reltype == RT.SLIDE_LAYOUT:
            self._copy_slide_layout_part(cast(SlideLayoutPart, source_part))
        else:
            self._copy_part_and_dependencies(source_part)

    def _copy_slide_layout_part(self, source_layout_part: SlideLayoutPart) -> None:
        """Copy a slide-layout, its slide-master, and the master's other layouts."""
        if source_layout_part in self._part_copies:
            return

        # -- copying the master also copies every layout it defines, which includes
        # -- this layout; this layout is reached again when the master's relationships
        # -- are reconstructed, so no extra wiring is needed here --
        source_master_part = cast(
            SlideMasterPart, source_layout_part.part_related_by(RT.SLIDE_MASTER)
        )
        self._copy_slide_master_part(source_master_part)

    def _copy_slide_master_part(self, source_master_part: SlideMasterPart) -> SlideMasterPart:
        """Copy a slide-master together with its theme and all of its slide-layouts."""
        existing = self._part_copies.get(source_master_part)
        if existing is not None:
            return cast(SlideMasterPart, existing)

        copied_master_part = cast(SlideMasterPart, self._construct_copy(source_master_part))
        self._part_copies[source_master_part] = copied_master_part

        # -- copy every sibling layout the master relates to and the rest of the
        # -- master's dependencies (theme, etc.); each layout in turn pulls in its
        # -- own dependencies (other than the already-copied master) --
        for rel in source_master_part.rels.values():
            if rel.is_external:
                continue
            if rel.reltype == RT.SLIDE_LAYOUT:
                source_layout_part = cast(SlideLayoutPart, rel.target_part)
                if source_layout_part not in self._part_copies:
                    copied_layout_part = cast(
                        SlideLayoutPart, self._construct_copy(source_layout_part)
                    )
                    self._part_copies[source_layout_part] = copied_layout_part
                    self._copy_dependencies_of(
                        source_layout_part, skip_reltypes={RT.SLIDE_MASTER}
                    )
            else:
                self._copy_part_and_dependencies(rel.target_part)

        return copied_master_part

    def _copy_dependencies_of(
        self, source_part: Part, skip_reltypes: set[str] | None = None
    ) -> None:
        """Recursively copy every part related to `source_part` (except skipped types)."""
        for rel in source_part.rels.values():
            if rel.is_external:
                continue
            if skip_reltypes and rel.reltype in skip_reltypes:
                continue
            if rel.reltype == RT.SLIDE:
                # -- a layout/master never legitimately links to a slide; if a linked
                # -- part is reached here it is a slide and must be selected --
                self._copy_slide_part(cast(SlidePart, rel.target_part))
            else:
                self._copy_part_and_dependencies(rel.target_part)

    def _copy_part_and_dependencies(self, source_part: Part) -> Part:
        """Copy an arbitrary part (image, chart, theme, ...) and its dependencies."""
        existing = self._part_copies.get(source_part)
        if existing is not None:
            return existing

        copied_part = self._construct_copy(source_part)
        self._part_copies[source_part] = copied_part
        self._copy_dependencies_of(source_part)
        return copied_part

    def _construct_copy(self, source_part: Part) -> Part:
        """Return a new, not-yet-related part in the destination package.

        The copy gets a fresh, collision-free partname. XML parts receive a deep copy of
        their element so the source package and a second import can never share mutable
        content; binary parts receive a copy of their blob.
        """
        reserved_partname = self._reserved_slide_partnames.get(source_part)
        partname = reserved_partname or self._next_partname_like(source_part.partname)

        if isinstance(source_part, XmlPart):
            return type(source_part)(
                partname, source_part.content_type, self._dst_package, source_part.clone_element()
            )

        return type(source_part)(
            partname, source_part.content_type, self._dst_package, source_part.blob
        )

    def _next_partname_like(self, source_partname: PackURI) -> PackURI:
        """Return an unused destination partname following the source partname pattern."""
        directory = source_partname.baseURI
        stem, ext = posixpath.splitext(source_partname.filename)

        match = re.match(r"^(.*?)(\d+)$", stem)
        # -- singleton-style partnames (no trailing index) keep their stem; on clash a
        # -- numeric suffix is appended --
        prefix = match.group(1) if match else stem

        tmpl = posixpath.join(directory, "%s%%d%s" % (prefix, ext))
        used_partnames = self._used_partnames

        for n in range(len(used_partnames) + 1, 0, -1):
            candidate = tmpl % n
            if candidate not in used_partnames:
                self._allocated_partnames.add(candidate)
                return PackURI(candidate)

        raise Exception("ProgrammingError: ran out of candidate partnames")  # pragma: no cover

    @property
    def _used_partnames(self) -> set[str]:
        return (
            {str(p.partname) for p in self._dst_package.iter_parts()}
            | self._allocated_partnames
        )

    # =================================================================================
    # phase 2 -- relationships
    # =================================================================================

    def _copy_relationships(self, source_part: Part, copied_part: Part) -> None:
        """Recreate the relationships of `source_part` on `copied_part` with the same rIds."""
        for rId, rel in source_part.rels.items():
            if rel.is_external:
                copied_part.rels.add_with_rId(rId, rel.reltype, rel.target_ref, is_external=True)
                continue

            # -- notes slides are not copied; nothing in the slide XML refers to them --
            if rel.reltype == RT.NOTES_SLIDE:
                continue

            target_copy = self._part_copies[rel.target_part]
            copied_part.rels.add_with_rId(rId, rel.reltype, target_copy)

    # =================================================================================
    # phase 3 -- presentation registration
    # =================================================================================

    def _register_slide_masters(self) -> None:
        """Relate each copied slide-master to the destination presentation part."""
        for source_part, copied_part in self._part_copies.items():
            if isinstance(source_part, SlideMasterPart):
                self._dst_pres_part.add_slide_master(cast(SlideMasterPart, copied_part))

    def _register_slides(self) -> tuple[SlidePart, ...]:
        """Relate the new slides to the presentation part in the requested order."""
        slide_parts: list[SlidePart] = []
        for source_slide_part in self._source_slide_parts:
            copied_slide_part = self._slide_copies[source_slide_part]
            self._dst_pres_part.add_slide_part(copied_slide_part)
            slide_parts.append(copied_slide_part)
        return tuple(slide_parts)

    @property
    def _selected_source_slides(self) -> set[SlidePart]:
        return set(self._source_slide_parts)
