import pydantic
from enum import Enum
from typing import List, Union

class Type(Enum):
    image = "image"
    text = "text"

class SlideItem(pydantic.BaseModel):
    id: str
    slide_number: int
    content: str
    type: Type
    order_number: int

class Image(SlideItem):
    image_bytes: bytes
    extension: str

    def metadata(self):
        return {
            "type": Type.image.value,
            "extension": self.extension,
            "image_id": self.id,
            "slide_number": self.slide_number,
            "order_number": self.order_number
        }

class Text(SlideItem):
    def metadata(self):
        return {
            "type": Type.text.value,
            "slide_number": self.slide_number,
            "order_number": self.order_number
        }

class Slide(pydantic.BaseModel):
    id: str
    slide_number: int
    items: List[Union[Image, Text]]


class Presentation(pydantic.BaseModel):
    id: str
    name: str
    slides: List[Slide]


class RAG_quizzer(pydantic.BaseModel):
    model_config = pydantic.ConfigDict(arbitrary_types_allowed=True)
    
    id: str
    name: str
    presentation: Presentation
    collection_id: str


