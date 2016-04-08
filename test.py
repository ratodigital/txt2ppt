#-*- coding: utf-8 -*-
from txt2ppt import Slides
import re

#
#for i in p.finditer('Tem um **bold** aqui e outro bold do outro lado. Mais um *bold* **inside** bem aqui ja e o terceiro'):
#    print i.start()
#    print i.group()

slides = Slides("slidex.pptx")
slides.new("Esse e um **Test** de negrito e tem *italico* tambem. Mais um outro *item realcado*, e **outro com** bold*")
slides.new("NOVO TEXTO")

slides.save()