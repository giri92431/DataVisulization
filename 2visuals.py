import win32com.client, MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)
import pandas as pd
import numpy as np 
import os



# Application = win32com.client.Dispatch("PowerPoint.Application")
# Application.Visible = True
# Presentation = Application.Presentations.Add()
# Base  = Presentation.Slides.Add(1, ppLayoutBlank)

# button_alpha = Base.Shapes.AddShape(msoShapeRectangle, 400, 10, 150, 40)
# button_alpha.TextFrame.TextRange.Text = 'Alphabetical'

# button_rating = Base.Shapes.AddShape(msoShapeRectangle, 560, 10, 150, 40)
# button_rating.TextFrame.TextRange.Text = 'By rating'


def animate(seq,image,trigger,path,duration=1.5):
	''' move image along the path when trigger is clicked'''
	effect = seq.AddEffect(
		Shape = image,
		effectId =msoAnimEffectPathDiamond,
		trigger=msoAnimTriggerOnShapeClick,)
	ani = effect.Behaviors.Add(msoAnimTypeMotion)
	ani.MotionEffect.Path = path
	effect.Timing.TriggerType = msoAnimTriggerWithPrevious
	effect.Timing.TriggerShape = trigger
	effect.Timing.Duration = duration

# seq_alpha = Base.TimeLine.InteractiveSequences.Add()
# seq_rating = Base.TimeLine.InteractiveSequences.Add()

df = pd.read_csv("name.csv")
df.columns = ["id","title"]

movies_alpha = sorted(df["title"], key=lambda v: v)
index_alpha = dict([(movie, i) for i , movie in enumerate(movies_alpha)])


df = df.as_matrix()
print(df)
# width, height = 28, 41
# for _id , _name in data:
# 	x = 10 + width * (_id % 25)
# 	y = 100 + height * (_id // 25)
# 	rect = Base.Shapes.AddShape(
#             msoShapeRectangle,
#             x, y,
#             width, height)
# 	rect.TextFrame.TextRange.Text = _name
# 	rect.TextFrame.TextRange.Font.Size = 6

# 	# alphabetical
# 	index = index_alpha[_name]
# 	animate(seq_alpha,rect,trigger= button_alpha,path ='M0,0 L{:.3f},{:.3f}'.format(
# 		((10 + width *(index % 25)- x )/750.),
# 		((100 + width *(index // 25) -y )/ 540.)
# 		))

# 	#by rating 
	
# 	animate(seq_rating,rect,trigger= button_rating,path ='M{:.3f},{:.3f} L0,0'.format(
# 		((10 + width *(index % 25)- x )/750.),
# 		((100 + width *(index // 25) -y )/ 540.)
# 		))
	
# 	print(index_alpha[_name] , _name)