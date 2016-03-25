#Instructions

1. Place git_diff_xls.py file in a folder

2. Create .gitconfig
3. Add the following line to the global .gitconfig:
'''python
[diff "zip"]
binary = True
textconv = python c:/path/to/git_diff_xlsx.py
'''
4. Create .gitattributes
5. Add the following line to the repository's .gitattributes
   *.xls diff=zip

6. After changing the xls file we can see changes on the local prompt as well as on repo master
7. Now, typing [git diff] at the prompt will produce text versions
   of Excel .xls files
8. "git status" will show diff of master repo.
