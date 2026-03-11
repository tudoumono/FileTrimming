from .base import BaseStep
from .step1_copy import Step1Copy
from .step2_normalize import Step2Normalize
from .step3_split import Step3Split
from .step4_markdown import Step4Markdown
from .step5_structure import Step5Structure
from .step6_chunk import Step6Chunk

ALL_STEPS: list[type[BaseStep]] = [
    Step1Copy,
    Step2Normalize,
    Step3Split,
    Step4Markdown,
    Step5Structure,
    Step6Chunk,
]

__all__ = ["BaseStep", "ALL_STEPS"]
