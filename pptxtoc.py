import argparse
import tempfile
import zipfile
import pdb ### Remove
import platform
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE


def getnotes(pptxfile):
    tmpd = tempfile.mkdtemp()
    zipfile.ZipFile(pptxfile).extractall(path=tmpd, pwd=None)

    # Parse notes content
    path = tmpd + '/ppt/notesSlides/'
    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        dom = parse(infile)
        noteslist = dom.getElementsByTagName('a:t')

        # The page number is part of the filename
        page = int(re.sub(r'\D', "", infile.split("/")[-1]))

        text = ''
        for node in noteslist:
            xmlTag = node.toxml()
            xmlData = xmlTag.replace('<a:t>', '').replace('</a:t>', '')
            text += " " + xmlData

        # Convert to ascii to simplify
        text = text.encode('ascii', 'ignore')
        words[str(page)] = text

    # Remove all the files created with unzip
    shutil.rmtree(tmpd)
    return words

if __name__ == "__main__":

    # Establish a default for the location of fonts to make it easier to specify the location
    # of the font.
    deffontdir = '' # Works OK for Windows
    if platform.system() == 'Darwin':
        deffontdir='/Library/Fonts/Microsoft/'
    elif platform.system() == 'Linux':
        # Linux stores fonts in sub-directories, so users must specify sub-dir and font name
        # beneath this directory.
        deffontdir='/usr/share/fonts/truetype/'

    # Parse command-line arguments
    parser = argparse.ArgumentParser(
            description='Create a Table of Content from a PowerPoint file',
            prog='pptxtoc.py',
            formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-f', action="store", dest="font", default='Tahoma',
            help="font directory")
    parser.add_argument('-F', action="store", dest="fontdir", default=deffontdir,
            help="font directory")
    parser.add_argument('-s', action="store", dest="stylepptx", default="Style.pptx",
            help="PowerPoint style document with no slides")
    parser.add_argument('-z', action="store", dest="fontsize", type=int, default=18,
            help="font size in points/pt")
    parser.add_argument('pptxfile', nargs=1)
    args = parser.parse_args()

    # Add the filename extension to font, relieving the user of this burden
    args.font += '.ttf'

    # Make sure the files exist and are readable
    if not (os.path.isfile(args.pptxfile) and os.access(args.pptxfile, os.R_OK)):
        sys.stderr.write("Cannot read the PowerPoint file \'%s\'.\n"%args.pptxfile)
        sys.exit(1)
    if not (os.path.isfile(args.stylepptx) and os.access(args.stylepptx, os.R_OK)):
        sys.stderr.write("Cannot read the PowerPoint style file \'%s\'.\n"%args.stylepptx)
        sys.exit(1)

    words = getnotes(args.pptxfile)
    pdb.set_trace()
    sys.exit(0)
    
    prs = Presentation("Style.pptx")
    blank_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(blank_slide_layout)
    
    slide.shapes.title.text = "Table of Contents"
    
    TEXT1="This is text inside a textbox\n"
    TEXT2="And a newline.\n"
    TEXT3="Another newline.\n"
    MAXDOTS=110
    
    top=Inches(1.75)
    left=Inches(.5)
    width=Inches(8)
    height=Inches(5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.auto_size=None
    tf.text = TEXT1+TEXT2+TEXT3
    p=tf.paragraphs[0]
    p.font.name = 'Tahoma'
    p.font.size=Pt(18)
    
    width=Inches(8.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = "."*(MAXDOTS-int(round((len(TEXT1)*1.5)))) + "\n" + \
                "."*(MAXDOTS-int(round((len(TEXT2)*1.5)))) + "\n" + \
                "."*(MAXDOTS-int(round((len(TEXT3)*1.5)))) + "\n" 
    tf.auto_size=None
    p=tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = 'Tahoma'
    p.font.size=Pt(18)
    
    
    left=left+Inches(8.5)
    width=Inches(.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.auto_size=None
    tf.text = "11\n13\n134\n"
    p=tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = 'Tahoma'
    p.font.size=Pt(18)
    
    
    prs.save('test.pptx')
