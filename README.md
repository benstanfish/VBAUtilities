# VBAUtilties

Personal collection of VBA code modules.

## News

Coming shortly: I'll add my graph-styles (including *SquareChart*) modules as well as my *Concrete PM Diagram* modules.

### Color Utilities

The following modules contain several functions related to the use of color in VBA contexts.

##### webcolors
Module containing all (140) W3 webcolor color names as Public Const as **Long** datatypes.

##### colorfuncs
Module containing functions that convert between rgb, hex, hsb and long as well as several contrast functions. Note that the module does contain a function for splitting rgb triplet strings into arrays, but it rgb triplet-like strings are used as the primary argument. Conversions of ranges to arrays from Excel --> VBA are a royal nightmare. This simplifies the process. Where rgb arrays are necessary, they are split once inside the VBA environment.
