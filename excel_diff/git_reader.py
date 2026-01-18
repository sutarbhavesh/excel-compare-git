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
        # Define the flag to hide the console window
        CREATE_NO_WINDOW = 0x08000000 

        normalized_path = path.replace("\\", "/")
        filename = f"[Branch: {branch}] {os.path.basename(normalized_path)}"
        save_path = os.path.join(target_dir, filename)

        if not url or not url.strip():
            repo_path = os.getcwd()
            cmd = ["git", "show", f"{branch}:{normalized_path}"]
            # Added creationflags here
            process = subprocess.run(
                cmd, cwd=repo_path, stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE, check=True,
                creationflags=CREATE_NO_WINDOW
            )
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
            # Added creationflags here
            subprocess.run(
                clone_cmd, check=True, capture_output=True,
                creationflags=CREATE_NO_WINDOW
            )

            show_cmd = ["git", "show", f"{branch}:{normalized_path}"]
            # Added creationflags here
            show_proc = subprocess.run(
                show_cmd, cwd=temp_dir, stdout=subprocess.PIPE, 
                check=True, creationflags=CREATE_NO_WINDOW
            )

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

    @staticmethod
    def fetch_commit_history(branch: str, path: str, url: Optional[str] = None, limit: int = 20) -> list:
        """Fetch commit history for a specific file."""
        CREATE_NO_WINDOW = 0x08000000
        normalized_path = path.replace("\\", "/")
        commits = []
        
        # Normalize branch name to lowercase to avoid case-sensitivity issues (e.g., 'Main' vs 'main')
        branch = branch.lower().strip() if branch else 'main'
        
        try:
            if not url or not url.strip():
                # Local repository
                repo_path = os.getcwd()
                # Use -- to specify the file path (this limits commits to just this file)
                cmd = ["git", "log", branch, "--pretty=format:%H|%an|%ai|%s", "--", normalized_path]
                process = subprocess.run(
                    cmd, cwd=repo_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    creationflags=CREATE_NO_WINDOW, text=True
                )
                
                if process.returncode != 0:
                    # Second try: without branch, just the path
                    cmd = ["git", "log", "--pretty=format:%H|%an|%ai|%s", "--", normalized_path]
                    process = subprocess.run(
                        cmd, cwd=repo_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        creationflags=CREATE_NO_WINDOW, text=True
                    )
                
                if process.returncode != 0:
                    raise RuntimeError(f"Git error: {process.stderr}")
            else:
                # Remote repository
                unique_id = int(time.time() * 1000) 
                temp_dir = os.path.abspath(os.path.join(os.environ.get('TEMP', '/tmp'), f"git_log_{unique_id}")).replace("\\", "/")
                
                try:
                    # Clone without --depth to get full history
                    clone_cmd = ["git", "clone", "--filter=blob:none", "--no-checkout", url, temp_dir]
                    clone_result = subprocess.run(clone_cmd, capture_output=True, creationflags=CREATE_NO_WINDOW, text=True)
                    
                    if clone_result.returncode != 0:
                        raise RuntimeError(f"Clone failed: {clone_result.stderr}")
                    
                    # Fetch the specific branch to ensure we have all commits (use normalized lowercase branch)
                    fetch_cmd = ["git", "fetch", "origin", branch]
                    subprocess.run(fetch_cmd, cwd=temp_dir, capture_output=True, creationflags=CREATE_NO_WINDOW, text=True)
                    
                    # Log for the specific file on the branch (-- limits to just this file)
                    cmd = ["git", "log", f"origin/{branch}", "--pretty=format:%H|%an|%ai|%s", "--", normalized_path]
                    process = subprocess.run(
                        cmd, cwd=temp_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        creationflags=CREATE_NO_WINDOW, text=True
                    )
                    
                    if process.returncode != 0:
                        raise RuntimeError(f"Log failed: {process.stderr}")
                finally:
                    if os.path.exists(temp_dir):
                        time.sleep(0.5)
                        try:
                            shutil.rmtree(temp_dir, onerror=GitReader._handle_remove_readonly)
                        except Exception:
                            pass

            if process.stdout:
                lines = process.stdout.strip().split('\n')
                for i, line in enumerate(lines[:limit]):
                    if line.strip():
                        parts = line.split('|', 3)
                        if len(parts) == 4:
                            commits.append({
                                'hash': parts[0][:7],
                                'full_hash': parts[0],
                                'author': parts[1],
                                'date': parts[2],
                                'message': parts[3]
                            })
        except Exception as e:
            raise RuntimeError(f"Error fetching commit history: {str(e)}")
        
        return commits
    @staticmethod
    def fetch_excel_by_commit(commit_hash: str, path: str, target_dir: str, url: Optional[str] = None) -> str:
        """Fetch Excel file from a specific commit."""
        CREATE_NO_WINDOW = 0x08000000
        normalized_path = path.replace("\\", "/")
        filename = f"[Commit: {commit_hash[:7]}] {os.path.basename(normalized_path)}"
        save_path = os.path.join(target_dir, filename)

        try:
            if not url or not url.strip():
                repo_path = os.getcwd()
                cmd = ["git", "show", f"{commit_hash}:{normalized_path}"]
                process = subprocess.run(
                    cmd, cwd=repo_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    creationflags=CREATE_NO_WINDOW, check=True
                )
                with open(save_path, "wb") as f:
                    f.write(process.stdout)
                return save_path
            else:
                unique_id = int(time.time() * 1000)
                temp_dir = os.path.abspath(os.path.join(target_dir, f"tmp_{unique_id}")).replace("\\", "/")
                
                try:
                    clone_cmd = ["git", "clone", "--depth", "1", "--filter=blob:none", "--no-checkout", url, temp_dir]
                    subprocess.run(clone_cmd, check=True, capture_output=True, creationflags=CREATE_NO_WINDOW)
                    
                    show_cmd = ["git", "show", f"{commit_hash}:{normalized_path}"]
                    show_proc = subprocess.run(
                        show_cmd, cwd=temp_dir, stdout=subprocess.PIPE, check=True,
                        creationflags=CREATE_NO_WINDOW
                    )
                    with open(save_path, "wb") as f:
                        f.write(show_proc.stdout)
                    return save_path
                finally:
                    if os.path.exists(temp_dir):
                        time.sleep(1)
                        try:
                            shutil.rmtree(temp_dir, onerror=GitReader._handle_remove_readonly)
                        except Exception:
                            pass
        except Exception as e:
            raise RuntimeError(f"Error fetching commit version: {str(e)}")