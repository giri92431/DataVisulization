import win32com.client
from MSO import constants
# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")

# content = new constants
# Add a presentation
Presentation = Application.Presentations.Add()

# Add a slide with a blank layout (12 stands for blank layout)
Base = Presentation.Slides.Add(1, 12)

# Add an oval. Shape 9 is an oval.
# oval = Base.Shapes.AddShape(constants.msoShapeFlowchartOr, 500, 100, 50, 100)