"""Copy selected slides, and all their package dependencies, between presentations."""

from __future__ import annotations

import posixpath
from copy import deepcopy
from typing import TYPE_CHECKING, Iterable, Iterator, cast

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.parts.image import ImagePart
from pptx.parts.slide import NotesMasterPart, NotesSlidePart, SlideMasterPart, SlidePart

if TYPE_CHECKING:
    from pptx.opc.package import _Relationship  # pyright: ignore[reportPrivateUsage]
    from pptx.package import Package
    from pptx.slide import Slide, Slides


class SlideImporter:
    """Copy selected slides from one open presentation into another.

    A single importer instance represents one independent import. Parts required by the selected
    slides (layouts, masters, images and other embedded resources) are deep-copied into the target
    package, so the target never references source-package parts or files on disk.
    """

    def __init__(self, target_slides: Slides):
        self._target_slides = target_slides
        self._target_package = target_slides.part.package
        self._copied_parts: dict[Part, Part] = {}
        self._rid_maps: dict[Part, dict[str, str]] = {}
        self._copied_master_parts: list[SlideMasterPart] = []

    def import_slides(self, slides: Iterable[Slide]) -> tuple[Slide, ...]:
        """Copy `slides` to the end of the target presentation and return the new slides."""
        source_slide_parts = self._validate_selection(slides)
        selected_slide_parts = set(source_slide_parts)
        self._validate_internal_links(source_slide_parts, selected_slide_parts)

        # Create disconnected slide stubs first so selected-slide-to-selected-slide relationships
        # can be remapped regardless of the order supplied by the caller.
        for source_slide_part in source_slide_parts:
            self._clone_part(source_slide_part)

        try:
            for source_slide_part in source_slide_parts:
                self._copy_dependencies(source_slide_part, selected_slide_parts)

            for source_part in tuple(self._copied_parts):
                self._copy_relationships(source_part, selected_slide_parts)

            self._rewrite_rIds()

            for master_part in self._copied_master_parts:
                self._commit_master(master_part)
            return tuple(
                self._commit_slide(source_slide_part)
                for source_slide_part in source_slide_parts
            )
        except Exception:
            # Nothing is attached to the target presentation until commit begins. Any error before
            # that point therefore leaves the target exactly as it was; avoid retaining detached
            # parts in package-level caches or mappings for a possible retry.
            self._copied_parts.clear()
            self._rid_maps.clear()
            self._copied_master_parts.clear()
            raise

    def _clone_part(self, source_part: Part) -> Part:
        """Return a detached deep copy of `source_part` in the target package."""
        existing_part = self._copied_parts.get(source_part)
        if existing_part is not None:
            return existing_part

        partname = self._next_partname(source_part.partname)
        if isinstance(source_part, XmlPart):
            target_part = type(source_part)(
                partname,
                source_part.content_type,
                self._target_package,
                deepcopy(cast("XmlPart", source_part)._element),
            )
        elif isinstance(source_part, ImagePart):
            target_part = ImagePart(
                partname,
                source_part.content_type,
                self._target_package,
                source_part.blob,
                source_part._filename,
            )
        else:
            target_part = type(source_part)(
                partname, source_part.content_type, self._target_package, source_part.blob
            )

        self._copied_parts[source_part] = target_part
        return target_part

    def _commit_master(self, source_master_part: SlideMasterPart) -> str:
        """Attach a copied slide-master part to the target presentation part."""
        target_master_part = cast(
            "SlideMasterPart", self._copied_parts[source_master_part]
        )
        target_presentation = self._target_slides.part.presentation_part
        rId = target_presentation.relate_to(target_master_part, RT.SLIDE_MASTER)
        target_presentation._element.get_or_add_sldMasterIdLst().add_sldMasterId(rId)
        return rId

    def _commit_slide(self, source_slide_part: SlidePart) -> Slide:
        """Attach a copied slide part to the target presentation in the requested order."""
        target_slide_part = cast("SlidePart", self._copied_parts[source_slide_part])
        presentation_part = self._target_slides.part
        rId = presentation_part.relate_to(target_slide_part, RT.SLIDE)
        self._target_slides._sldIdLst.add_sldId(rId)
        return target_slide_part.slide

    def _copy_dependencies(
        self, source_part: Part, selected_slide_parts: set[SlidePart]
    ) -> None:
        """Recursively copy the non-slide relationship graph needed by `source_part`."""
        for relationship in source_part.rels.values():
            if relationship.is_external:
                continue

            target_part = relationship.target_part
            if isinstance(target_part, NotesMasterPart):
                # A copied notes slide uses the target presentation's existing notes master. Avoid
                # importing a second package-level notes master, which PowerPoint does not expect.
                continue
            if relationship.reltype == RT.SLIDE:
                if (
                    not isinstance(source_part, NotesSlidePart)
                    and target_part not in selected_slide_parts
                ):
                    # This was rejected during pre-copy validation when the relationship was used.
                    # Skip unused slide relationships instead of cloning an unselected slide.
                    continue
                # Selected slide stubs already exist; notes back-references are copied later.
                continue

            if target_part in self._copied_parts:
                continue

            target_copy = self._clone_part(target_part)
            if isinstance(target_copy, SlideMasterPart):
                self._copied_master_parts.append(cast("SlideMasterPart", target_part))
            self._copy_dependencies(target_part, selected_slide_parts)

    def _copy_relationships(
        self, source_part: Part, selected_slide_parts: set[SlidePart]
    ) -> None:
        """Copy relationships from one source part to its copied counterpart."""
        target_part = self._copied_parts[source_part]
        rid_map = self._rid_maps.setdefault(source_part, {})

        for relationship in source_part.rels.values():
            if relationship.is_external:
                new_rId = target_part.relate_to(
                    relationship.target_ref, relationship.reltype, is_external=True
                )
            else:
                related_source_part = relationship.target_part
                if isinstance(related_source_part, NotesMasterPart):
                    related_target_part = self._target_slides.part.notes_master_part
                elif relationship.reltype == RT.SLIDE:
                    related_target_part = self._copied_parts.get(related_source_part)
                    if related_target_part is None:
                        # Selected slide relationships are in the copy map. An unselected slide
                        # relationship can only be an unused relationship left by another client.
                        continue
                else:
                    related_target_part = self._copied_parts[related_source_part]

                new_rId = target_part.relate_to(related_target_part, relationship.reltype)

            rid_map[relationship.rId] = new_rId

    def _next_partname(self, source_partname: PackURI) -> PackURI:
        """Return an unused target partname derived from `source_partname`."""
        filename = source_partname.filename
        stem, ext = posixpath.splitext(filename)
        match = PackURI._filename_re.match(stem)
        if match is not None and match.group(2):
            name = match.group(1)
            index_length = len(match.group(2))
            suffix = stem[len(name) + index_length :] + ext
            template = posixpath.join(source_partname.baseURI, f"{name}%d{suffix}")
        else:
            template = posixpath.join(source_partname.baseURI, f"{stem}%d{ext}")

        for candidate in self._iter_candidate_partnames(template):
            if candidate not in self._used_partnames:
                self._used_partnames.add(candidate)
                return candidate
        raise AssertionError("could not allocate target partname")  # pragma: no cover

    def _iter_candidate_partnames(self, template: str) -> Iterator[PackURI]:
        """Generate partnames using `template`, including gaps left by deleted parts."""
        for number in range(1, len(self._used_partnames) + 2):
            yield PackURI(template % number)

    def _relationship_is_used(
        self, source_part: Part, relationship: _Relationship
    ) -> bool:
        """Return whether a slide relationship backs a hyperlink in `source_part`."""
        if not isinstance(source_part, XmlPart):
            return True
        rId = relationship.rId
        for element in source_part._element.iter():
            if element.tag in _HLINK_TAGS and element.get(_R_ID_ATTR) == rId:
                return True
        return False

    def _rewrite_rIds(self) -> None:
        """Rewrite relationship ids in every copied XML part belonging to this import."""
        for copied_source_part, target_part in tuple(self._copied_parts.items()):
            if not isinstance(target_part, XmlPart):
                continue
            rid_map = self._rid_maps.get(copied_source_part)
            if not rid_map:
                continue
            for element in target_part._element.iter():
                old_rId = element.get(_R_ID_ATTR)
                if old_rId in rid_map:
                    element.set(_R_ID_ATTR, rid_map[old_rId])

    def _validate_internal_links(
        self, slide_parts: list[SlidePart], selected_slide_parts: set[SlidePart]
    ) -> None:
        """Raise before target mutation if a copied hyperlink points outside the selection."""
        visited: set[Part] = set()

        def visit(part: Part):
            if part in visited:
                return
            visited.add(part)
            for relationship in part.rels.values():
                if relationship.is_external:
                    continue
                target_part = relationship.target_part
                if isinstance(target_part, NotesMasterPart):
                    continue
                if relationship.reltype == RT.SLIDE:
                    if (
                        target_part not in selected_slide_parts
                        and self._relationship_is_used(part, relationship)
                    ):
                        raise ValueError(
                            "cannot import slide because it links to a slide that was not selected"
                        )
                    if target_part not in selected_slide_parts:
                        continue
                visit(target_part)

        for slide_part in slide_parts:
            visit(slide_part)

    def _validate_selection(self, slides: Iterable[Slide]) -> list[SlidePart]:
        """Validate the caller supplied slides from one other, already-open presentation."""
        source_slide_parts: list[SlidePart] = []
        source_package: Package | None = None

        for slide in slides:
            source_part = slide.part
            if not isinstance(source_part, SlidePart):
                raise TypeError("all selected items must be Slide objects")
            if source_part.package is self._target_package:
                raise ValueError("cannot import a slide that already belongs to this presentation")
            if source_package is None:
                source_package = source_part.package
            elif source_part.package is not source_package:
                raise ValueError("all selected slides must belong to the same presentation")
            if source_part in source_slide_parts:
                raise ValueError("a slide may only be selected once in a single import")
            source_slide_parts.append(source_part)

        return source_slide_parts

    @property
    def _used_partnames(self) -> set[PackURI]:
        used_partnames = self.__dict__.get("_used_partnames_value")
        if used_partnames is None:
            used_partnames = {part.partname for part in self._target_package.iter_parts()}
            self.__dict__["_used_partnames_value"] = used_partnames
        return cast("set[PackURI]", used_partnames)


_HLINK_TAGS = frozenset((qn("a:hlinkClick"), qn("a:hlinkHover")))
_R_ID_ATTR = qn("r:id")
