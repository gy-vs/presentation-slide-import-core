# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.slideimporter` module."""

from __future__ import annotations

import pytest

from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.slide import SlidePart
from pptx.slideimporter import SlideImporter

from .unitutil.mock import instance_mock, loose_mock


class DescribeSlideImporter:
    """Unit-test suite for `pptx.slideimporter.SlideImporter` objects."""

    def it_allocates_unique_partnames_preserving_names_and_extensions(self):
        importer = SlideImporter.__new__(SlideImporter)
        importer.__dict__["_used_partnames_value"] = {
            PackURI("/ppt/media/image1.png"),
            PackURI("/ppt/slides/slide1.xml"),
        }

        image_partname = importer._next_partname(PackURI("/ppt/media/image1.png"))
        slide_partname = importer._next_partname(PackURI("/ppt/slides/slide1.xml"))
        second_slide_partname = importer._next_partname(PackURI("/ppt/slides/slide1.xml"))

        assert image_partname == PackURI("/ppt/media/image2.png")
        assert slide_partname == PackURI("/ppt/slides/slide2.xml")
        assert second_slide_partname == PackURI("/ppt/slides/slide3.xml")

    def it_rejects_a_slide_that_already_belongs_to_the_target(self, request):
        target_package_ = instance_mock(request, Package)
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.package = target_package_
        slide_ = loose_mock(request, name="slide")
        slide_.part = slide_part_
        importer = self._importer(request, target_package_)

        with pytest.raises(ValueError, match="already belongs"):
            importer._validate_selection([slide_])

    def it_rejects_slides_from_different_source_presentations(self, request):
        target_package_ = instance_mock(request, Package)
        source_package_1_ = instance_mock(request, Package)
        source_package_2_ = instance_mock(request, Package)
        slide_part_1_ = instance_mock(request, SlidePart)
        slide_part_1_.package = source_package_1_
        slide_part_2_ = instance_mock(request, SlidePart)
        slide_part_2_.package = source_package_2_
        slide_1_ = loose_mock(request, name="slide-1")
        slide_1_.part = slide_part_1_
        slide_2_ = loose_mock(request, name="slide-2")
        slide_2_.part = slide_part_2_
        importer = self._importer(request, target_package_)

        with pytest.raises(ValueError, match="same presentation"):
            importer._validate_selection([slide_1_, slide_2_])

    def _importer(self, request, package_):
        target_slides_ = loose_mock(request, name="target-slides")
        target_slides_.part.package = package_
        return SlideImporter(target_slides_)
