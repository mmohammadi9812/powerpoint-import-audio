import os, sys, shutil, re, subprocess, zipfile, logging
from typing import List, Tuple, Callable, Union, Any, Optional
from pathlib import Path

import tkinter as tk
from tkinter import filedialog


logging.basicConfig(level=logging.DEBUG,
    filename=Path(Path(__file__).parent / 'log.txt'),
    format='%(asctime)s - %(levelname)s - %(lineno)d - %(message)s')


AUDIO_TARGET_NAME = 'audio'
AUDIO_TARGET_GLOB = 'ppt/media/*.m4a'


def ffmpeg() -> str:
    if shutil.which('ffmpeg') is None or (sys.platform == 'win32' and shutil.which('ffmpeg.exe') is None):
        raise ValueError('ffmpeg doesn\'t exists!')

    return 'ffmpeg.exe' if sys.platform == 'win32' else 'ffmpeg'


def natural_sort(l: List[str]) -> List[Any]:
    convert = lambda text: int(text) if text.isdigit() else text.lower()
    alphanum_key = lambda key: [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(l, key=alphanum_key)


def get_files(initdir: Optional[str]=None, filetypes: Tuple[Tuple[str]]=(("PowerPoint", "*.pptx"),) ,title: str="Choose powerpoint files.") -> Union[Tuple[str], str]:
    root = tk.Tk()
    root.withdraw()
    initdir = initdir if initdir is not None else os.getcwd()
    pptx_files = filedialog.askopenfilenames(
        initialdir=initdir,
        filetypes=filetypes,
        title=title)
    return pptx_files


def get_target(filepath: Union[str, Path]) -> Path:
    filepath = Path(filepath) if type(filepath) is str else filepath
    name = filepath.stem
    parent = filepath.parent
    logging.debug(f'name: {name}')
    logging.debug(f'parent: {parent}')
    target = Path(parent / name)
    return target

get_audio_target: Union[Callable[[str], Path], Callable[[Path], Path]] = lambda path: Path(get_target(path) / AUDIO_TARGET_NAME)

def extract_audio(filepath: Union[str, Path]) -> None:
    target = get_target(filepath)
    logging.debug(f'target: {target}')
    target.mkdir()
    with zipfile.ZipFile(filepath, 'r') as zip:
        zip.extractall(target)


def prune(filepath: Union[str, Path]) -> None:
    target = get_target(filepath)
    audio_target = get_audio_target(filepath)
    audio_target.mkdir()
    audio_files = target.glob(AUDIO_TARGET_GLOB)
    for audio in audio_files:
        audio_name = Path(audio).name
        logging.debug(f'audio: {audio}')
        shutil.move(audio, Path(audio_target / audio_name))
    useless = os.listdir(target)
    useless.remove('audio')
    for f in useless:
        path = Path(target / f)
        logging.debug(f'removing {path} ...')
        if path.is_dir():
            shutil.rmtree(path)
        else:
            path.unlink()


def convert_to_pcm(filepath: Union[str, Path]) -> List[str]:
    audio_target = get_audio_target(filepath)
    audio_files = audio_target.glob('*.m4a')
    for file in audio_files:
        file_name = Path(file).stem
        logging.debug(f'converting {file} ...')
        process = subprocess.run([command, '-i', file, '-c:a', 'pcm_s16le', '-ac', '2', '-ar', '48000', '-f', 's16le', f'{file_name}.pcm'], check=True, stdout=subprocess.PIPE)
        logging.debug(process.stdout)
        Path(file).unlink()
    pcm_files = natural_sort([path.name for path in audio_target.glob('*.pcm')])
    logging.debug(f'pcm_files: {pcm_files}')
    return pcm_files


def append_pcm_files(filepath: Union[str, Path], pcmfiles: List[str]) -> Path:
    audio_target = get_audio_target(filepath)
    outpcm = Path(audio_target / 'out.pcm')
    with open(outpcm, 'ab') as out:
        for pcm in pcm_files:
            with open(pcm, 'rb') as p:
                out.write(p.read())
    for pcm in pcm_files:
        pcm.unlink()
    return outpcm


def out_pcm_to_m4a(filepath: Union[str, Path], outpcm: Path, outm4a: str='out.m4a') -> Path:
    audio_target = get_audio_target(filepath)
    command = ffmpeg()
    os.chdir(audio_target)
    process = subprocess.run([command, '-f', 's16le', '-ac', '2', '-ar', '48000', '-i', f'{outpcm}', '-c:a', 'aac', '-b:a', '192k', '-ac', '2', outm4a], check=True, stdout=subprocess.PIPE)
    logging.debug(process.stdout)
    logging.debug(process.stderr)
    outpcm.unlink()
    return Path(audio_target / outm4a)


if __name__ == "__main__":
    logging.debug(sys.platform)
    cwd = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    command = ffmpeg()
    if not Path(cwd).exists():
        raise ValueError(f'{cwd} doesn\'t exists!')
    cwd = Path(cwd)
    logging.debug(f'cwd: {cwd}')
    # pptx_files = cwd.glob("*.pptx")
    # logging.debug(f'pptx_files: {pptx_files}')

    pptx_files = get_files()
    logging.debug(f'file_path: #{pptx_files}# type: {type(pptx_files)}')

    for pptx in pptx_files:
        name = Path(pptx).stem
        parent = Path(pptx).parent
        logging.debug(f'parent: {parent}')
        logging.debug(f'name: {name}')
        target = Path(parent / name)
        extract_audio(pptx)
        audio_target = Path(target / 'audio')
        prune(pptx)
        os.chdir(audio_target)
        # DONE: merge audio files to one file
        pcm_files = convert_to_pcm(pptx)
        outpcm = append_pcm_files(pptx, pcm_files)
        outm4a = out_pcm_to_m4a(pptx, outpcm)
        os.chdir(parent)
        shutil.move(Path(audio_target / outm4a), Path(parent / (name + '.m4a')))
        shutil.rmtree(name)


