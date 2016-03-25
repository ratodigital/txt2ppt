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

import StringIO

from txt2ppt import Slides
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
    ('/ppt', Ppt),
], debug=True)
