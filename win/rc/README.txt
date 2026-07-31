Note the files twapi_events.h twapi_events.rc *.bin are generated
by the message compiler. Do not edit them as changes will be lost.
As to why these files are in the repository and not generated
during builds, see comment at top of twapi_events.man.

Generate files by running the command

     mc.exe -um -b -h . -r . -n twapi_events.man
