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

    def fetch_all_test_units(self) -> dict[str, TestUnit]:
        test_units = dict()
        for index in range(1, self.count + 1):
            tu_inst = self.item(index)
            test_units[tu_inst.name] = tu_inst
        return test_units

