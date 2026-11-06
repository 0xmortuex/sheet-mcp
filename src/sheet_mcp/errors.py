"""One exception type for every failure an agent can act on."""


class SheetError(Exception):
    """An expected failure whose message tells the agent what to do next."""
