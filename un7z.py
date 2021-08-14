# -*- coding: utf-8 -*-
"""使い方：python un7z.py --target=[ファイル名またはディレクトリ]
実行した後パスワードを入力することが必要
ファイルはxxx.7zのファイル、
ディレクトリの中でいくつかのxxx.7zファイルがあります（すべての.7zファイルは同じパスワードを使用）
＊clickとpy7zrのライブラリが必要
"""
import click
import os
from datetime import datetime
from py7zr import SevenZipFile


ext = '.7z'


def _extract(file):
    """extract contents of a filename.7z file
    and write to /filename directory
    """
    password = f'osaka{datetime.now().strftime("%y%m")}'
    try:
        path = os.path.abspath(file)
        out_dir = f'{os.path.split(path)[0]}/{file[:file.index(ext)]}'
        with SevenZipFile(file, 'r', password=password) as archive:
            archive.extractall(path=out_dir)
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
            _extract(target)
        except Exception as e:
            print(f'An error occurred while '
                  f'extracting contents of {target}: {e}')
    else:
        try:
            for file in os.listdir(target):
                if file.endswith(ext):
                    _extract(file, password)
        except Exception as e:
            print(f'An error occurred while reading files in {target}: {e}')


if __name__ == '__main__':
    extract()

