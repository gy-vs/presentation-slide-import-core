"""Unit-test suite for `pptx.slideimporter` module."""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx import Presentation
from pptx.enum.action import PP_ACTION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.image import ImagePart
from pptx.slideimporter import SlideImportError


class Describe_SlideImporter(object):
    """Unit-test suite for slide-import behavior across real packages."""

    def it_deep_copies_text_picture_layout_and_external_hyperlink(self, two_slide_source):
        dst = Presentation()
        src = two_slide_source

        new_slides = dst.slides.import_slides([src.slides[0], src.slides[1]])

        slide = new_slides[0]
        assert slide.shapes.title.text == "Quarterly Report"
        picture = next(s for s in slide.shapes if s.shape_type == 13)
        assert picture.click_action.action == PP_ACTION.HYPERLINK
        assert picture.click_action.hyperlink.address == "https://example.com/report"
        # -- the slide now depends on its own copied layout and image parts --
        layout_part = slide.part.part_related_by(RT.SLIDE_LAYOUT)
        assert layout_part.package is dst.part.package
        image_part = self._image_part(slide)
        assert isinstance(image_part, ImagePart)
        assert image_part.package is dst.part.package

    def it_redirects_an_internal_jump_to_the_imported_copy_when_target_is_selected(
        self, two_slide_source
    ):
        dst = Presentation()
        src = two_slide_source

        s1, s2 = dst.slides.import_slides([src.slides[0], src.slides[1]])

        action = s1.shapes.title.click_action
        assert action.action == PP_ACTION.NAMED_SLIDE
        assert action.target_slide is s2
        # -- the slide relationship points at a part in the destination package --
        target_part = self._related_slide_part(s1)
        assert target_part is s2.part
        assert target_part.package is dst.part.package

    def it_follows_the_requested_order_for_internal_jumps(self, two_slide_source):
        dst = Presentation()
        src = two_slide_source

        appendix, title_slide = dst.slides.import_slides([src.slides[1], src.slides[0]])

        assert title_slide.shapes.title.click_action.target_slide is appendix

    def it_raises_when_an_internal_jump_target_was_not_selected(self, two_slide_source):
        dst = Presentation()
        src = two_slide_source
        slide_count_before = len(dst.slides)

        with pytest.raises(SlideImportError):
            dst.slides.import_slides([src.slides[0]])

        # -- failure leaves the destination usable and unchanged --
        assert len(dst.slides) == slide_count_before
        stream = io.BytesIO()
        dst.save(stream)
        Presentation(io.BytesIO(stream.getvalue()))

    def it_appends_imported_slides_after_existing_ones(self, two_slide_source):
        dst = Presentation()
        original = dst.slides.add_slide(dst.slide_layouts[0])
        src = two_slide_source

        new_slides = dst.slides.import_slides([src.slides[0], src.slides[1]])

        all_slides = list(dst.slides)
        assert all_slides[0] is original
        assert all_slides[1:] == list(new_slides)

    def it_leaves_the_source_presentation_unchanged(self, two_slide_source):
        src = two_slide_source
        src_slide1_partname = src.slides[0].part.partname
        src_image_part = src.slides[0].part.related_part(
            next(
                rId
                for rId, rel in src.slides[0].part.rels.items()
                if rel.reltype == RT.IMAGE
            )
        )
        src_image_bytes = src_image_part.blob

        dst = Presentation()
        dst.slides.import_slides([src.slides[0], src.slides[1]])

        assert len(src.slides) == 2
        assert src.slides[0].part.partname == src_slide1_partname
        picture = next(s for s in src.slides[0].shapes if s.shape_type == 13)
        assert picture.click_action.hyperlink.address == "https://example.com/report"
        assert src.slides[0].shapes.title.click_action.target_slide is src.slides[1]
        assert src_image_part.blob == src_image_bytes
        # -- source still saves and reopens --
        stream = io.BytesIO()
        src.save(stream)
        reopened = Presentation(io.BytesIO(stream.getvalue()))
        assert len(reopened.slides) == 2

    def it_produces_independent_copies_on_repeated_import(self, two_slide_source):
        dst = Presentation()
        src = two_slide_source

        first = dst.slides.import_slides([src.slides[0], src.slides[1]])
        second = dst.slides.import_slides([src.slides[0], src.slides[1]])

        assert first[0].part is not second[0].part
        # -- mutating the first import does not affect the second --
        first[0].shapes.title.text_frame.text = "CHANGED"
        assert second[0].shapes.title.text == "Quarterly Report"
        # -- distinct image parts with distinct partnames --
        first_image = self._image_part(first[0])
        second_image = self._image_part(second[0])
        assert first_image is not second_image
        assert first_image.partname != second_image.partname

    def it_keeps_same_named_image_parts_from_different_packages_distinct(
        self, request, two_slide_source
    ):
        from PIL import Image as PIL_Image

        def deck_with_image(color):
            image_bytes = io.BytesIO()
            PIL_Image.new("RGB", (20, 20), color).save(image_bytes, format="PNG")
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(
                io.BytesIO(image_bytes.getvalue()), 0, 0
            )
            return prs

        dst = Presentation()
        deck_a = deck_with_image((200, 10, 10))
        deck_b = deck_with_image((10, 10, 200))

        dst.slides.import_slides([deck_a.slides[0]])
        dst.slides.import_slides([deck_b.slides[0]])

        stream = io.BytesIO()
        dst.save(stream)
        reopened = Presentation(io.BytesIO(stream.getvalue()))
        blobs = {
            self._image_part(slide).blob
            for slide in reopened.slides
            if any(sh.shape_type == 13 for sh in slide.shapes)
        }
        assert len(blobs) == 2

    def it_produces_a_self_contained_package_that_reopens(self, two_slide_source):
        dst = Presentation()
        dst.slides.add_slide(dst.slide_layouts[0])
        src = two_slide_source
        dst.slides.import_slides([src.slides[0], src.slides[1]])

        stream = io.BytesIO()
        dst.save(stream)
        data = stream.getvalue()
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            rels_bytes = b"".join(
                zf.read(n) for n in zf.namelist() if n.endswith(".rels")
            )
        assert b"https://example.com/report" in rels_bytes

        reopened = Presentation(io.BytesIO(data))
        assert len(reopened.slides) == 3
        title_slide, appendix = reopened.slides[1], reopened.slides[2]
        assert title_slide.shapes.title.text == "Quarterly Report"
        picture = next(s for s in title_slide.shapes if s.shape_type == 13)
        assert picture.click_action.hyperlink.address == "https://example.com/report"
        assert title_slide.shapes.title.click_action.target_slide is appendix
        assert "Appendix" in appendix.shapes[0].text_frame.text

    def it_validates_its_arguments(self, two_slide_source):
        dst = Presentation()
        src = two_slide_source

        with pytest.raises(SlideImportError):
            dst.slides.import_slides([])

        with pytest.raises(SlideImportError):
            dst.slides.import_slides([src.slides[0], src.slides[0]])

        with pytest.raises(SlideImportError):
            dst.slides.import_slides([dst.slides.add_slide(dst.slide_layouts[0])])

    # -- helpers -----------------------------------------------------

    def _related_slide_part(self, slide):
        return next(
            slide.part.related_part(rId)
            for rId, rel in slide.part.rels.items()
            if rel.reltype == RT.SLIDE
        )

    def _image_part(self, slide):
        return next(
            slide.part.related_part(rId)
            for rId, rel in slide.part.rels.items()
            if rel.reltype == RT.IMAGE
        )

    # fixtures ------------------------------------------------------

    @pytest.fixture
    def two_slide_source(self):
        from pptx.util import Inches

        prs = Presentation()
        title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        title_slide.shapes.title.text = "Quarterly Report"
        picture = title_slide.shapes.add_picture(
            "tests/test_files/python-icon.jpeg", Inches(1), Inches(2), Inches(2), Inches(2)
        )
        picture.click_action.hyperlink.address = "https://example.com/report"

        appendix = prs.slides.add_slide(prs.slide_layouts[6])
        appendix.shapes.add_textbox(
            Inches(1), Inches(1), Inches(4), Inches(1)
        ).text_frame.text = "Appendix"

        title_slide.shapes.title.click_action.target_slide = appendix
        return prs
