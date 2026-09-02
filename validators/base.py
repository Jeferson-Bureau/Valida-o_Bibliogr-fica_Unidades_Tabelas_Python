from abc import ABC, abstractmethod
from typing import Optional
from models import ReferenceInput, ValidationResult

class BaseValidator(ABC):
    def __init__(self):
        self.api_name = "Base"

    @abstractmethod
    def validate(self, reference: ReferenceInput) -> Optional[ValidationResult]:
        """Recebe uma referência de entrada e retorna o resultado da validação, ou None se não encontrar."""
        pass
