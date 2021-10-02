# -*- coding: utf-8 -*-
"""使い方：python un7z.py --target=[ファイル名またはディレクトリ]
ファイルはxxx.7zのファイル、
ディレクトリの中でいくつかのxxx.7zファイルがあります（すべての.7zファイルは同じパスワードを使用）
＊clickとpy7zrのライブラリが必要
"""
import click
import os
from datetime import datetime
from py7zr import SevenZipFile
from zipfile import ZipFile


EXT_7Z = '.7z'
EXT_ZIP = '.zip'


def _extract(file):
    """extract contents of a filename.7z file
    and write to /filename directory
    """
    password = f'tpe{datetime.now().strftime("%Y%m")}'
    out_dir, ext = os.path.splitext(file)
    try:
        # out_dir = file[:file.index(ext)]
        if ext == EXT_7Z:
            with SevenZipFile(file, 'r', password=password) as archive:
                archive.extractall(path=out_dir)
        elif ext == EXT_ZIP:
            with ZipFile(file, "r") as archive:
                for info in archive.infolist():
                    info.filename = info.filename.encode('cp437')\
                        .decode('cp932')
                archive.extractall(path=out_dir, pwd=password.encode('utf-8'))
        print(f'Extracted contents of {file} to {out_dir}')
    except Exception as e:
        print(f'An error occurred while extracting contents of {file}: {e}')


@click.command()
@click.option('--target', default=None,
    help='File or directory with 7z files to decompress.')
def extract(target):
    """extract contents of a .7z file or
    .7z files in a directory (all files have the same password)
    """
    if target is None:
        print('No file or directory specified: --target=file/directory')
        return
    if os.path.isfile(target):
        try:
            _extract(os.path.abspath(target))
        except Exception as e:
            print(f'An error occurred while '
                  f'extracting contents of {target}: {e}')
    else:
        try:
            for file in os.listdir(target):
                if file.endswith(EXT_7Z) or file.endswith(EXT_ZIP):
                    _extract(os.path.abspath(f'{target}/{file}'))
        except Exception as e:
            print(f'An error occurred while reading files in {target}: {e}')


if __name__ == '__main__':
    extract()
