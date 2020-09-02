class Point(object):
    def __init__(self, description, x, y):
        self.description = description
        self.x = x
        self.y = y

    def __repr__(self):
        return "{0}: (x={1}, y={2})".format(self.description, self.x, self.y)