class Treemap:
    def __init__(self, width, height, data, draw):
        '''Treemap(width, height, data, fn) '''
        self.x, self.y = 0.0, 0.0
        self.scale  = (float(width * height) / self.get(data, sum)) ** 0.5
        self.width  = float(width)  / self.scale
        self.height = float(height) / self.scale
        self.draw   = draw
        self.squarify(data, [], min(self.width, self.height))
        
    def get(self, data, fn): return fn((x[1] for x in data))

    def layoutrow(self, row):
        if self.width >= self.height:
            dx = self.get(row, sum) / self.height
            step = self.height / len(row)
            for i,v in enumerate(row): 
                self.draw(self.scale * self.x, self.scale * (self.y + i * step), self.scale * dx, self.scale * step, v)
            self.x += dx
            self.width -= dx
        else:
            dy = self.get(row, sum) / self.width
            step = self.width / len(row)
            for i,v in enumerate(row):
                self.draw(self.scale * (self.x + i * step), self.scale * self.y, self.scale * step, self.scale * dy, v)
            self.y += dy
            self.height -= dy

    def aspect(self, row, w):
        s = self.get(row, sum)
        return max(w*w*self.get(row, max)/s/s, s*s/w/w/self.get(row, max))

    def squarify(self, children, row, w):
        if not children.any():
            if row: self.layoutrow(row)
            return
        c = children[0]
        if not row or self.aspect(row, w) > self.aspect(row + [c], w):
            self.squarify(children[1:], row + [c], w)
        else:
            self.layoutrow(row)
            self.squarify(children, [], min(self.height, self.width))

