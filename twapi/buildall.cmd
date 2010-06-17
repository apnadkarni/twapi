rem Builds all configurations of TWAPI
nmake %*
nmake distribution %*
nmake tmdistribution  %*
nmake SERVERONLY=1  %*
nmake distribution SERVERONLY=1  %*
nmake tmdistribution SERVERONLY=1  %*
nmake DESKTOPONLY=1  %*
nmake distribution DESKTOPONLY=1  %*
nmake tmdistribution DESKTOPONLY=1  %*
