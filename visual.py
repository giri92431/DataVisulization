import win32com.client, MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)
import pandas as pd
import numpy as np 
import os



Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
Presentation = Application.Presentations.Add()
Base  = Presentation.Slides.Add(1, ppLayoutBlank)

button_alpha = Base.Shapes.AddShape(msoShapeRectangle, 400, 10, 150, 40)
button_alpha.TextFrame.TextRange.Text = 'Alphabetical'

button_rating = Base.Shapes.AddShape(msoShapeRectangle, 560, 10, 150, 40)
button_rating.TextFrame.TextRange.Text = 'By rating'


df = pd.read_csv("name.csv")
df.columns = ["id","title"]
movies_alpha = sorted(df["title"], key=lambda v: v)
index_alpha = dict([(movie, i) for i , movie in enumerate(movies_alpha)])


df = df.as_matrix()


def fileName(movie_id):
	return os.path.join("D:\\DataVisulization\\imdb\\"+ str(movie_id +1 )+ ".jpeg")



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

seq_alpha = Base.TimeLine.InteractiveSequences.Add()
seq_rating = Base.TimeLine.InteractiveSequences.Add()

width ,height = 30,41
for _id , _name in df:
	x = 10 + width * (_id % 25)
	y = 100 + height * (_id // 25)
	print(x,y)
	image = Base.Shapes.AddPicture(os.path.abspath(fileName(_id)),LinkToFile=True,SaveWithDocument=False,Left =x ,Top =y, Width=width,Height=height)
	link = image.ActionSettings(ppMouseClick).Hyperlink
	link.Address = "http://www.imdb.com/chart/top"
	link.ScreenTip =  str(_name)

	#alphabetical
	index =index_alpha[_name]
	animate(seq_alpha,image,trigger= button_alpha,path ='M0,0 L{:.3f},{:.3f}'.format(
		((10 + width *(index % 25)- x )/750.),
		((100 + width *(index // 25) -y )/ 540.)
		))

	#by rating 
	
	animate(seq_rating,image,trigger= button_rating,path ='M{:.3f},{:.3f} L0,0'.format(
		((10 + width *(index % 25)- x )/750.),
		((100 + width *(index // 25) -y )/ 540.)
		))








	
