import win32com.client, MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)
import pandas as pd
import numpy as np 


Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
Presentation = Application.Presentations.Add()
Slide = Presentation.Slides.Add(1, ppLayoutBlank)

data = pd.read_csv('assets1.csv')
data = np.array(data)

# data['value'] =data['value'].apply(lambda x : int(x.replace(",","")))
# data.to_csv("test.csv")

from Treemap import Treemap
def draw(x, y, w, h, n):
    shape = Slide.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    shape.TextFrame.TextRange.Text = n[0] + ' (' + str(int(n[1]/1000 + 500 )) + 'M)'
    shape.TextFrame.MarginLeft = shape.TextFrame.MarginRight = 0
    shape.TextFrame.TextRange.Font.Size = 11


Treemap(720, 540, data, draw)
