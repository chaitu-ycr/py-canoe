import win32com.client

from py_canoe.core.child_elements.test_unit import TestUnit


class TestUnits:
    """The TestUnits object represents the collection of test units of a single test configuration."""

    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> TestUnit:
        return TestUnit(self.com_object.Item(index))

    def add(self, path: str) -> TestUnit:
        return TestUnit(self.com_object.Add(path))

    def remove(self, index: int):
        self.com_object.Remove(index)

