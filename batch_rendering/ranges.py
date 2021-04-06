

class Range:
    """All the range are inclusive"""

    def __init__(self, num):
        self.rows = list(range(num))

    def between_rows(self, start, end):
        "return a list from one row to another row specified."
        return self.rows[start-2:end-1]

    def from_row(self, row):
        "return a list from row specified."
        return self.rows[row-2:]

    def to_row(self, row):
        "return a list to row specified."
        return self.rows[:row-1]

    def unique_row(self, row):
        "return a list of single element."
        return self.rows[row-2:row-1]
