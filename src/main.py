import docx2txt
import os
import pathlib
import re
import sys
import argparse
from PIL import Image

def convert_to_gif(img, gif_location):
    im = Image.open(img)
    gif_name = img.stem + ".gif"
    gif_file= pathlib.Path(gif_location,gif_name)
    im.save(gif_file)



text = """Izvlacenje slika iz .docx fajla, -br za broj propisa/programa. npr:
python main.py -br 1594401
da bi se dobile slike: 1594401-01.png itd. za dalju obradu.
"""

parser = argparse.ArgumentParser(description=text)
parser.add_argument("-br", "--brojPropisa", help="Broj propisa/programa", action="append")
args = parser.parse_args()

#width = 450-590
#height = 150-200
if len(sys.argv)==1:
    # display help message when no args are passed.
    parser.print_help()
    sys.exit(1)

broj_propisa = (args.brojPropisa)[0]

current_location = os.getcwd()
cwd_location = pathlib.Path(current_location)

wordFiles_location = pathlib.Path(cwd_location.parent,"WordFiles")
if(not wordFiles_location.exists()):
    print("kreiramo poctno okruzenje")
    wordFiles_location.mkdir()
    sys.exit(status="Ubaciti .docx fajl u WordFiles folder i probajte ponovo")

images_location =  pathlib.Path(cwd_location.parent, "images")
if(not images_location.exists()):
    print("kreiram images folder gde ce biti slike izvucene iz fajla")
    images_location.mkdir()

gif_location = pathlib.Path(cwd_location.parent, "GIFs")
if(not gif_location.exists()):
    gif_location.mkdir()

flist=[]
for p in wordFiles_location.iterdir():
    if p.is_file():
        flist.append(p)


for f in flist:
    print(f)

delete_files = []
try:
    for p in images_location.iterdir():
        if p.is_file():
            delete_files.append(p)
    for p in delete_files:
        p.unlink()
except OSError as e:
    if(e.errno != e.ENOENT):
        raise
    print("obrisati slike rucno pa probati ponovo")
    pass

try:
    for t in gif_location.iterdir():
        if t.is_file():
            t.unlink()
except OSError as e:
    if(e.errno != e.ENOENT):
        raise
    print("obrisati sve slike iz GIF foldera")
    pass

print("brisem slike od prethodne obrade")

print("izvlacim slike....")
text = docx2txt.process(flist[0], images_location)

print(f'renamovanje fajlova {broj_propisa}')
#15440601- ime programa
all_images = []
renamed_images = []
for img in images_location.iterdir():
    if img.is_file():
        all_images.append(img.stem)
        image_file = img.stem
        old_number = re.match(r'.*?([0-9]+)\.', img.name).group(1)
        if(len(old_number) == 1):
            old_number = "0"+old_number
        old_extension = img.suffix
        old_dir = img.parent
        new_name = broj_propisa +"-" + old_number + old_extension
        img.rename(pathlib.Path(old_dir, new_name))
        renamed_images.append(img)


print("konvertovanje u GIF...")
for img in images_location.iterdir():
    convert_to_gif(img, gif_location)


print("zavrseno")