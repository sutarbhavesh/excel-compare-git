import subprocess
import os
import shutil
import time
import stat
from typing import Optional

class GitReader:
    @staticmethod
    def _handle_remove_readonly(func, path, excinfo):
        """Forces the deletion of read-only files (standard in .git folders)."""
        os.chmod(path, stat.S_IWRITE)
        func(path)

    @staticmethod
    def fetch_excel(branch: str, path: str, target_dir: str, url: Optional[str] = None) -> str:
        normalized_path = path.replace("\\", "/")
        filename = f"[Branch: {branch}] {os.path.basename(normalized_path)}"
        save_path = os.path.join(target_dir, filename)

        if not url or not url.strip():
            repo_path = os.getcwd()
            cmd = ["git", "show", f"{branch}:{normalized_path}"]
            process = subprocess.run(cmd, cwd=repo_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
            with open(save_path, "wb") as f:
                f.write(process.stdout)
            return save_path

        unique_id = int(time.time() * 1000) 
        temp_dir = os.path.abspath(os.path.join(target_dir, f"tmp_{unique_id}")).replace("\\", "/")
        
        try:
            clone_cmd = [
                "git", "clone", "--depth", "1", "--filter=blob:none", 
                "--no-checkout", url, temp_dir
            ]
            subprocess.run(clone_cmd, check=True, capture_output=True)

            show_cmd = ["git", "show", f"{branch}:{normalized_path}"]
            show_proc = subprocess.run(show_cmd, cwd=temp_dir, stdout=subprocess.PIPE, check=True)

            with open(save_path, "wb") as f:
                f.write(show_proc.stdout)

            return save_path

        except Exception as e:
            raise RuntimeError(f"Error accessing private repo: {str(e)}")
        
        finally:
            if os.path.exists(temp_dir):
                time.sleep(1) 
                try:
                    shutil.rmtree(temp_dir, onerror=GitReader._handle_remove_readonly)
                except Exception:
                    print(f"Cleanup deferred for {temp_dir}")