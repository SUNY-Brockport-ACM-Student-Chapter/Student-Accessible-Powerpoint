"""
Smoke tests for the Pydantic data model.

These guard invariant #1 (order_number exists and is an int) at the type
level, so a reckless rename of the field fails a test rather than silently
misalignining alt text at runtime.
"""

from __future__ import annotations

import pytest


def test_slide_item_requires_order_number():
    from models.models import Text, Type

    with pytest.raises(Exception):
        # order_number intentionally omitted; pydantic must reject this.
        Text(id="1", slide_number=1, content="hi", type=Type.text)


def test_slide_item_order_number_is_int():
    from models.models import Text, Type

    with pytest.raises(Exception):
        Text(
            id="1",
            slide_number=1,
            content="hi",
            type=Type.text,
            order_number="not-an-int",  # type: ignore[arg-type]
        )


def test_image_carries_bytes_and_extension():
    from models.models import Image, Type

    img = Image(
        id="i1",
        slide_number=1,
        content="",
        type=Type.image,
        order_number=0,
        image_bytes=b"\x89PNG\r\n",
        extension="png",
    )
    md = img.metadata()
    # metadata must preserve order_number for the join invariant.
    assert md["order_number"] == 0
    assert md["slide_number"] == 1
    assert md["type"] == "image"


def test_presentation_slide_item_union_roundtrip():
    from models.models import Image, Presentation, Slide, Text, Type

    deck = Presentation(
        id="p1",
        name="deck.pptx",
        slides=[
            Slide(
                id="s1",
                slide_number=1,
                items=[
                    Text(id="t1", slide_number=1, content="hello", type=Type.text, order_number=0),
                    Image(
                        id="i1",
                        slide_number=1,
                        content="",
                        type=Type.image,
                        order_number=1,
                        image_bytes=b"x",
                        extension="png",
                    ),
                ],
            )
        ],
    )
    assert [it.order_number for it in deck.slides[0].items] == [0, 1]
