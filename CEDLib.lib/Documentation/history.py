# -*- coding: utf-8 -*-
"""Small browser-style page and heading navigation history."""


class NavigationHistory(object):
    def __init__(self):
        self._items = []
        self._position = -1

    @property
    def can_back(self):
        return self._position > 0

    @property
    def can_forward(self):
        return 0 <= self._position < len(self._items) - 1

    @property
    def current(self):
        if 0 <= self._position < len(self._items):
            return self._items[self._position]
        return None

    def push(self, path, anchor=""):
        item = (str(path), str(anchor or ""))
        if self.current == item:
            return item
        if self.can_forward:
            self._items = self._items[: self._position + 1]
        self._items.append(item)
        self._position = len(self._items) - 1
        return item

    def back(self):
        if not self.can_back:
            return None
        self._position -= 1
        return self.current

    def forward(self):
        if not self.can_forward:
            return None
        self._position += 1
        return self.current

    def clear(self):
        self._items = []
        self._position = -1

