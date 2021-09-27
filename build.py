import os
import shutil
import tarfile
import time


pwd = os.getcwd()
script_folder = os.path.join(pwd, 'thermal')
program = 'thermal.py'
release_folder = os.path.join(pwd, 'release')
zip_folder = os.path.join(pwd, release_folder,time.strftime('%Y%m%d%H%M%S', time.localtime(time.time())))
temp_files = ('build', 'dist', 'thermal.spec', '__pycache__', 'thermal.py')

def copy_files():
    os.chdir(pwd)
    os.mkdir(zip_folder)
    shutil.copy(os.path.join('thermal', 'thermal.py'), os.path.join(zip_folder, 'thermal.py'))
    shutil.copy(os.path.join('doc', 'readme.pdf'), os.path.join(zip_folder, 'readme.pdf'))
    shutil.copytree('ref', os.path.join(zip_folder, 'templates'))

def make_exe():
    os.chdir(zip_folder)
    os.system('pyinstaller -F {0}'.format(program))
    shutil.move(os.path.join('dist', 'thermal.exe'), os.path.join('thermal.exe'))
    # delete tmp files
    for path in temp_files:
        if os.path.isfile(path):
            os.remove(path)
            continue
        shutil.rmtree(path, ignore_errors=True)

def make_targz():
    os.chdir(release_folder)
    with tarfile.open(os.listdir(release_folder)[-1] + '_thermal.tar.gz', "w:gz") as tar:
        tar.add(zip_folder, arcname=os.path.basename(zip_folder + '_thermal'))
    shutil.rmtree(zip_folder, ignore_errors=True)

def build():
    copy_files()
    make_exe()
    make_targz()

if __name__ == '__main__':
    build()