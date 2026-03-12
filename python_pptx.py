from pptx import Presentation
import sys


prs = Presentation( sys.argv[1] )

f = open( sys.argv[2], "a" )

#print(eachfile)
#print("----------------------")
for slide in prs.slides:
    f.write( "----------------------\n" )
    f.write( str( slide.slide_id ) )
    f.write( "\n----------------------\n" )
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            slide_text = shape.text.strip( )
            slide_text = slide_text.replace( "\n\n", "\n" )
            slide_text = slide_text.replace( "\n\n", "\n" )
            slide_text = slide_text.replace( "\n\n", "\n" )
            slide_text = slide_text.replace( "\n\n", "\n" )
            slide_text = slide_text.replace( "\n\n", "\n" )
            slide_text = slide_text.replace( "  ", " " )
            slide_text = slide_text.replace( "  ", " " )
            slide_text = slide_text.replace( "  ", " " )
            slide_text = slide_text.replace( "  ", " " )
            slide_text = slide_text.replace( "  ", " " )
            slide_text += "\n\n"
            f.write( slide_text )
            
f.close()
