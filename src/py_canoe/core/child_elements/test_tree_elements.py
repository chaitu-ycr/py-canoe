import win32com.client

from py_canoe.core.child_elements.test_tree_element import TestTreeElement


class TestTreeElements:
    """The TestTreeElements object represents a collection of test tree elements."""

    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def count(self) -> int:
        return self.com_object.Count

    def item(self, index: int) -> TestTreeElement:
        return TestTreeElement(self.com_object.Item(index))

    def add(self, name: str) -> TestTreeElement:
        return TestTreeElement(self.com_object.Add(name))

    def remove(self, index: int):
        self.com_object.Remove(index)

