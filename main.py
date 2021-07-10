import zipfile
import os
import sys
import shutil

def unzip_file(input_path,output_path):
    filename = input_path[input_path.rfind("\\") + 1:]
    if '.pptx' not in filename:
        print('Please make sure file is in .pptx format')
        return 0
    new_output_path = output_path+'\\temp_'+filename
    z = zipfile.ZipFile(input_path, 'r')
    os.makedirs(new_output_path, exist_ok=True)
    for file in z.namelist():
        z.extract(file, new_output_path)
    if not get_activex(new_output_path+'\\ppt\\activeX',output_path+'\\final_'+filename):
        print('There is no Flash (SWF) in this PowerPoint')
        shutil.rmtree(new_output_path)
        return 0
    shutil.rmtree(new_output_path)
    return 1

def get_activex(input_path,output_path):
    try:
        files = os.listdir(input_path)
    except Exception:
        return 0
    for file in files:
        if '.xml' in file and 'active' in file:
            with open(input_path+'\\'+file, 'r') as f:
                content = f.read()
            if 'D27CDB6E-AE6D-11CF-96B8-444553540000' in content:
                decode_activex(input_path+'\\'+file[:file.rfind('.')] + '.bin',output_path,file[:file.rfind('.')])
    return 1

def decode_activex(input_path,output_path,filename):
    with open(input_path,'rb') as f:
        content = f.read()
    position = content.find(b'\x46\x57\x53')
    s = content[position+4:position+8].hex()
    t = ''
    for x in range(len(s)-1,0,-2):
        t += s[x-1] + s[x]
    length = int(t,16)
    os.makedirs(output_path,exist_ok=True)
    with open(output_path+'\\'+filename+'.swf','wb') as f:
        f.write(content[position:position+length])

if __name__ == '__main__':
    while True:
        path = input()
        path = path.replace('"','')
        if unzip_file(path,os.path.abspath(os.path.dirname(sys.argv[0]))):
            print('Finish')

