#include <fstream.h>
#include <iostream.h>
#include <string.h>

#define MIME_LINE_SIZE  71

static
unsigned char bintoasc[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

void OutputChar(ostream & out, unsigned char c)
{
  static int i = 0;
  
  out << bintoasc[c];
  i++;
  if(i>MIME_LINE_SIZE)
    {
      out << "\n";
      i=0;
    }
}

int main(int argc, char *argv[])
{
    if(argc<3)
        {
        cout << "Usage:\n\tMIME <infile.bin> <outfile.asc>";
        return 0;
        }

    fstream infile(argv[1], ios::in | ios::binary);
    
    if(!infile)
        {
        cout << "Unable to open " << argv[1] << " for input!";
        return 0;
        }

    fstream outfile(argv[2], ios::out | ios::trunc);

    if(!outfile)
        {
        cout << "Unable to open " << argv[1] << " for output!";
        return 0;
        }

    unsigned char c1, c2, c3, c4;
    int b1, b2, b3;

    while(!infile.eof())
        {
	  b1 = infile.get();
	  b2 = infile.get();
	  b3 = infile.get();
	  
	  if(b1!=EOF)
	    {
	      c1 = (b1 & 0xFF) >> 2;
	      c2 = (b1 & 0x03) << 4;
	    }

	  if(b2!=EOF)
	    {
	      c2 = c2 | ((b2 & 0xF0) >> 4);
	      c3 = (b2 & 0x0F) << 2;
	    }
	  
	  if(b3!=EOF)
	    {
	      c3 = c3 | ((b3 & 0xC0) >> 6);
	      c4 = b3 & 0x3F;
	    }

	  if(b1!=EOF)
	    {
	      OutputChar(outfile, c1);
	      OutputChar(outfile, c2);
	    }

	  if(b2!=EOF)
	    {
	      OutputChar(outfile, c3);
	    }
	  
	  if(b3!=EOF)
	    {
	      OutputChar(outfile, c4);
	    }
        }
}
