import os
import subprocess
import platform
from pathlib import Path
import shutil
import signal
import psutil  # 需要安装 pip install psutil

from loguru import logger


class ConvertToPdfError(Exception):
    def __init__(self, msg):
        self.msg = msg
        super().__init__(self.msg)


def check_fonts_installed():
    """Check if required Chinese fonts are installed."""
    system_type = platform.system()

    if system_type in ['Windows', 'Darwin']:
        pass
    else:
        # Linux: use fc-list
        try:
            output = subprocess.check_output(['fc-list', ':lang=zh'], encoding='utf-8')
            if output.strip():  # 只要有任何输出（非空）
                return True
            else:
                logger.warning(
                    f"No Chinese fonts were detected, the converted document may not display Chinese content properly."
                )
        except Exception:
            pass


def get_soffice_command():
    """Return the path to LibreOffice's soffice executable depending on the platform."""
    system_type = platform.system()

    # First check if soffice is in PATH
    soffice_path = shutil.which('soffice')
    if soffice_path:
        return soffice_path

    if system_type == 'Windows':
        # Check common installation paths
        possible_paths = [
            Path(os.environ.get('PROGRAMFILES', 'C:/Program Files')) / 'LibreOffice/program/soffice.exe',
            Path(os.environ.get('PROGRAMFILES(X86)', 'C:/Program Files (x86)')) / 'LibreOffice/program/soffice.exe',
            Path('C:/Program Files/LibreOffice/program/soffice.exe'),
            Path('C:/Program Files (x86)/LibreOffice/program/soffice.exe')
        ]

        # Check other drives for windows
        for drive in ['C:', 'D:', 'E:', 'F:', 'G:', 'H:']:
            possible_paths.append(Path(f"{drive}/LibreOffice/program/soffice.exe"))

        for path in possible_paths:
            if path.exists():
                return str(path)

        raise ConvertToPdfError(
            "LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/ "
            "or ensure soffice.exe is in your PATH environment variable."
        )
    else:
        # For Linux/macOS, provide installation instructions if not found
        try:
            # Try to find soffice in standard locations
            possible_paths = [
                '/usr/bin/soffice',
                '/usr/local/bin/soffice',
                '/opt/libreoffice/program/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return path

            raise ConvertToPdfError(
                "LibreOffice not found. Please install it:\n"
                "  - Ubuntu/Debian: sudo apt-get install libreoffice\n"
                "  - CentOS/RHEL: sudo yum install libreoffice\n"
                "  - macOS: brew install libreoffice or download from https://www.libreoffice.org/\n"
                "  - Or ensure soffice is in your PATH environment variable."
            )
        except Exception as e:
            raise ConvertToPdfError(f"Error locating LibreOffice: {str(e)}")


def convert_file_to_pdf(input_path, output_dir):
    """Convert a single document (ppt, doc, etc.) to PDF."""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"The input file {input_path} does not exist.")

    os.makedirs(output_dir, exist_ok=True)

    check_fonts_installed()

    soffice_cmd = get_soffice_command()

    cmd = [
        soffice_cmd,
        '--headless',
        '--norestore',
        '--invisible',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(input_path)
    ]

    try:
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            stdout, stderr = process.communicate(timeout=60)
            if process.returncode != 0:
                logger.error(f"LibreOffice convert failed: {stderr.decode()}")
        except subprocess.TimeoutExpired:
            logger.error("LibreOffice convert timed out. Killing process tree...")
            kill_process_and_children(process.pid)
            logger.error("LibreOffice process tree killed due to timeout.")
    except Exception as e:
        logger.error(f"Failed to execute LibreOffice: {e}")

def kill_process_and_children(pid):
    """Kill a process and all its children using psutil."""
    try:
        parent = psutil.Process(pid)
        children = parent.children(recursive=True)
        for child in children:
            logger.warning(f"Killing child process PID {child.pid} ({child.name()})")
            child.kill()
        logger.warning(f"Killing parent process PID {parent.pid} ({parent.name()})")
        parent.kill()
    except psutil.NoSuchProcess:
        logger.warning(f"Process {pid} already exited.")
    except Exception as e:
        logger.error(f"Failed to kill process tree: {e}")
