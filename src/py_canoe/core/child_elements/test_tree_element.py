import win32com.client


class TestTreeElement:
    """The TestTreeElement object represents a single element in the test tree."""

    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    def __getattr__(self, item):
        return getattr(self.com_object, item)

    @property
    def caption(self) -> str:
        return self.com_object.Caption

    @property
    def enabled(self) -> bool:
        return self.com_object.Enabled

    @enabled.setter
    def enabled(self, value: bool):
        self.com_object.Enabled = value

    @property
    def id(self) -> str:
        return self.com_object.Id

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def type(self) -> int:
        return self.com_object.Type

    @property
    def elements(self):
        # Imported lazily to avoid circular imports
        from py_canoe.core.child_elements.test_tree_elements import TestTreeElements

        return TestTreeElements(self.com_object.Elements)

