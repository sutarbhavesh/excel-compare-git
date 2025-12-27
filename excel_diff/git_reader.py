import subprocess
import os
from typing import List, Optional


class GitReader:
    """
    Read-only helper for accessing files from a git repository.
    Uses git CLI internally (no third-party libraries).
    """

    def __init__(self, repo_path: str):
        if not os.path.isdir(repo_path):
            raise ValueError("Invalid repository path")

        self.repo_path = os.path.abspath(repo_path)

    def _run_git(self, args: List[str]) -> str:
        """
        Run a git command safely inside the repo.
        """
        completed = subprocess.run(
            ["git"] + args,
            cwd=self.repo_path,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip())

        return completed.stdout

    # ----------------------------
    # Branch operations
    # ----------------------------

    def list_branches(self) -> List[str]:
        """
        Returns a list of local branches.
        """
        output = self._run_git(["branch", "--list"])
        return [line.strip().lstrip("* ").strip() for line in output.splitlines()]

    def branch_exists(self, branch: str) -> bool:
        """
        Check if a branch exists.
        """
        try:
            self._run_git(["rev-parse", "--verify", branch])
            return True
        except RuntimeError:
            return False

    # ----------------------------
    # File operations
    # ----------------------------

    def file_exists(self, branch: str, file_path: str) -> bool:
        """
        Check if a file exists in a given branch.
        """
        try:
            self._run_git(["ls-tree", "-r", "--name-only", branch, file_path])
            return True
        except RuntimeError:
            return False

    def read_file(self, branch: str, file_path: str) -> bytes:
        """
        Read a file from a given branch.
        Returns raw bytes (important for Excel).
        """
        normalized_path = file_path.replace("\\", "/")

        output = subprocess.run(
            ["git", "show", f"{branch}:{normalized_path}"],
            cwd=self.repo_path,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if output.returncode != 0:
            raise RuntimeError(output.stderr.decode("utf-8", errors="ignore"))

        return output.stdout
