# Emails
Multiple projects dedicated to email handling.


Each project will be its own branch named 'master_branch-name' from solution 2 at [Stack Overflow](https://stackoverflow.com/questions/14679614/whats-the-best-practice-for-putting-multiple-projects-in-a-git-repository)

> repo 1
>
> git push origin master:master-1
> 
> repo 2
>
> git push origin master:master-2


## How To Add Another Project

### File structure

- Emails
  - .git
  - .gitignore
  - README.md
  - Project-1
    - trunk
      - .git
	  - Project-1 files
	- branch
  - Project-2
    - trunk
	  - Project-2 files


### Git Command Line

```git
cd "Project-2/trunk"
git init
git add .
git status
git commit -m 'commit message'
git remote add origin 'https://github.com/PeterGeras/Emails.git'
git push origin master:master_project-2
```