import subprocess
import os
import shutil
import time
import stat
import requests
from typing import List, Optional

class GitReader:
    """
    Fast Git helper using HTTP requests for remote files to handle 2GB+ repos.
    Falls back to local Git CLI for files already on the disk.
    """

    @staticmethod
    def _handle_remove_readonly(func, path, excinfo):
        os.chmod(path, stat.S_IWRITE)
        func(path)

    @staticmethod
    def fetch_excel(branch: str, path: str, target_dir: str, url: Optional[str] = None) -> str:
        normalized_path = path.replace("\\", "/")
        filename = f"git_{branch}_{os.path.basename(normalized_path)}"
        save_path = os.path.join(target_dir, filename)

        # --- CASE 1: REMOTE GITHUB (FAST REQUEST METHOD) ---
        if url and "github.com" in url:
            try:
                # Convert GitHub Web URL to Raw URL
                raw_url = url.replace("github.com", "raw.githubusercontent.com")
                if raw_url.endswith(".git"):
                    raw_url = raw_url[:-4]
                
                full_raw_path = f"{raw_url}/{branch}/{normalized_path}"

                # Download only the specific file
                response = requests.get(full_raw_path, timeout=15)
                
                if response.status_code == 200:
                    with open(save_path, "wb") as f:
                        f.write(response.content)
                    return save_path
                else:
                    raise RuntimeError(f"GitHub Error {response.status_code}: Could not find file at {full_raw_path}")

            except Exception as e:
                raise RuntimeError(f"Network Error: {str(e)}")

        # --- CASE 2: LOCAL REPOSITORY (GIT SHOW METHOD) ---
        else:
            repo_path = os.getcwd()
            try:
                cmd = ["git", "show", f"{branch}:{normalized_path}"]
                process = subprocess.run(cmd, cwd=repo_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                with open(save_path, "wb") as f:
                    f.write(process.stdout)
                return save_path
            except Exception as e:
                raise RuntimeError(f"Local Git Error: {str(e)}")