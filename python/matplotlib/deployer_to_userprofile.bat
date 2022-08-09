@echo off
set prog_path=%userprofile%/Documents/GitHub/deployer
set prog=%prog_path%/deployer.pl
set tgt_path=%userprofile%/.matplotlib/stylelib
set f1=pubfig.mplstyle

perl %prog% ^
--nopause ^
--path=%tgt_path% ^
%f1% ^
