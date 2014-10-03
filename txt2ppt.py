import sys
sys.path.insert(0, 'libs')

import os
import urllib

from google.appengine.api import users
#from google.appengine.ext import ndb

import jinja2
import webapp2

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)

DEFAULT_GUESTBOOK_NAME = 'admin'

import fix_path # has to be first.

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor

import lxml

import StringIO

#http://python-pptx.readthedocs.org/en/latest/user/text.html

class Slides:
	prs = Presentation()
    	file_name = "slides.pptx"
    	blank_slide_layout = prs.slide_layouts[6]
    	font_size = Pt(30)
    	font_color = "ffffff"

	def __init__(self, file_name):
        	self.prs = Presentation()
        	self.file_name = file_name	

	def set_font_size(self, size):
		self.font_size = Pt(size)

	def set_font_color(self, color):
		self.font_color = color
		
	def new(self, text):
		slide = self.prs.slides.add_slide(self.blank_slide_layout)

		left = Inches (0.5)
		top = Inches (2)
		width = Inches (9)
		height = Inches (3)
	
		txBox = slide.shapes.add_textbox(left, top, width, height)
		tf = txBox.textframe
		tf.word_wrap = True
		tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

		p = tf.add_paragraph()
		p.text = text
		p.font.size = self.font_size
		p.font.color.rgb = RGBColor.from_string(self.font_color);
		p.alignment = PP_ALIGN.CENTER

	def save(self):
		self.prs.save(self.file_name)

		
class MainPage(webapp2.RequestHandler):

    def get(self):
        guestbook_name = self.request.get('guestbook_name',
                                          DEFAULT_GUESTBOOK_NAME)
        #greetings_query = Greeting.query(
        #    ancestor=guestbook_key(guestbook_name)).order(-Greeting.date)
        #greetings = greetings_query.fetch(10)

        if users.get_current_user():
            url = users.create_logout_url(self.request.uri)
            url_linktext = 'Logout'
        else:
            url = users.create_login_url(self.request.uri)
            url_linktext = 'Login'

        template_values = {
            #'greetings': greetings,
            'guestbook_name': urllib.quote_plus(guestbook_name),
            'url': url,
            'url_linktext': url_linktext,
        }

        template = JINJA_ENVIRONMENT.get_template('index.html')
        self.response.write(template.render(template_values))



class Ppt(webapp2.RequestHandler):

    def post(self):
        username = self.request.get('username',
                                          DEFAULT_GUESTBOOK_NAME)

        fontsize = self.request.get('fontsize')
        fontcolor = self.request.get('fontcolor')
		
        slides = self.create_slides(self.request.get('content'), int(fontsize), fontcolor[1:])
        
        self.response.headers['Content-Type'] = 'application/octet-stream'
        self.response.headers['Content-Disposition'] = 'attachment; filename=slides.pptx'
        self.response.headers['Content-Length'] = len(slides);
        self.response.write(slides)

    def create_slides(self, text, fontsize, fontcolor):  
        output = StringIO.StringIO()
        s1 = Slides(output)
        s1.set_font_size(fontsize)
        s1.set_font_color(fontcolor)

        for line in text.split("\n"):
            if (line.strip()):
                s1.new(line.strip())

        s1.save()
        contents = output.getvalue()
        return contents;

application = webapp2.WSGIApplication([
    ('/', MainPage),
    #('/sign', Guestbook),
    ('/ppt', Ppt),
], debug=True)
