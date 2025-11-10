@rem Upgrade all Chocolatey packages
choco upgrade all -y --no-progress >choco-upgrade-%ComputerName%.txt
@rem Make a list of packages associated with this computer
choco list >choco-packages-%ComputerName%.txt
@rem Remove all the shortcuts placed on the desktop
del C:\Users\Public\Desktop\*.lnk
del C:\Users\Public\Desktop\*.url
C:\Users\bjs54\Miniconda3\envs\py314\python.exe package_updates.py check_updated
