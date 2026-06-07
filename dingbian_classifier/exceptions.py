"""Domain-specific exceptions for the classifier."""


class ClassifierError(Exception):
    """Base class for user-facing classifier errors."""


class TargetSheetNotFoundError(ClassifierError):
    """Raised when no production detail sheet matches the naming rule."""


class MissingRequiredFieldsError(ClassifierError):
    """Raised when the target sheet lacks required columns."""


class RuleConfigError(ClassifierError):
    """Raised when classification rules are invalid."""

