import git
import os

git_dir = os.getcwd() 
g = git.cmd.Git(git_dir)
g.pull()