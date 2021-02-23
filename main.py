import sys
import os
import docx2txt as d2t
import fire
import glob
import zipfile
from datetime import date
from PIL import Image

today = date.today()
d = today.strftime("%b-%d-%Y")
res = d2t.process("hi.docx", "./img")
boxes = ["Additional Information", "NB", "Definition", "Caution" ]
pwd = os.path.abspath(os.getcwd())
print("PWD ->>>" + pwd)

def start(zipf):
    '''
    Runs the entire Script -> Extracts zip and imgs from docxs and compresses
    imgs
    Parameters
    ----------
    zipf : string
        Zipfile that gets extracted

    Returns
    -------
    None
        Returns a new folder in which the extracted/compressed files can be
        found
    '''


    # make folder with name of zipfile
    foldername = "newsletter_" +  d
    destinationfolder = os.path.join(os.path.curdir, foldername)
    img_path = os.path.join(destinationfolder, "imgs")
    print("IMG PATH --> ", img_path)
    os.mkdir(destinationfolder)
    os.mkdir(img_path)

    # extract zip to that new folder (maybe move zipfile there too)
    extractfiles(zipf, destinationfolder)

    # extract images from docx
    extract_imgs(destinationfolder, img_path)

    # compress images that are now in that folder
    compressImages(img_path)


def extractfiles(z, dest):
    '''
    Extracts the Zipfile into the given destination
    Parameters
    ----------
    z : string
        Zipfile to extract
    dest: string
        Path to the folder to which files should get extracted

    Returns
    -------
    None
        Extraced files from zip in given destination
    '''
    with zipfile.ZipFile("./" + z, "r") as zip_ref:
        zip_ref.extractall(dest)

def extract_imgs(dest, imgs):
    '''
    Extract imges from Wordfiles in given destination folder
    Parameters
    ----------
    dest : string
        Path to the folder to from where the imgs get extraced
        
    imgs: string
        Path to the folder to which imgs should get saved

    Returns
    -------
    None
        Extraced files from zip in given destination
    '''
    wordfiles = []

    for file in os.listdir(dest):
        if file.endswith('.docx'):
            wordfiles.append(file)

    for i in range(len(wordfiles)):
        res = d2t.process(dest + "/" + wordfiles[i], imgs)

def compressImages(dest):
    '''
    Compress all images from given directory
    Parameters
    ----------
    dest : string
        Path to the folder where images are

    Returns
    -------
    None
        Compressed images
    '''
    image_list = []
    img_ext = ["jpg", "jpeg", "png"]
    # add all images to image-list
    for ext in img_ext:
        for filename in glob.glob(dest + f'/*.{ext}'): #assuming ext
            im=Image.open(filename)
            image_list.append(im)

    img_i = 0
    for im in image_list:
        img_i += 1
        if im.format in ["JPEG", "PNG", "JPG"]:
            newfile_name = im.filename[im.filename.rfind("/"):]
            print(newfile_name)
            im.save(dest + f"/new{img_i}.{im.format}", quality= 40)

if __name__ == '__main__':
    fire.Fire()
