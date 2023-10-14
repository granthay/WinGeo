# WinGeo
WinGeo

Description:
    WinGeo is a suite of windows programs that process geodetic data.

Architecture:
    All programs share code, all the shared code is in the folder CommonCode
    All programs are working towards a modified model view processor presenter framework see below.
        > Models - Classes that represent data. These include file classes, record classes, and      collection classes.
        > Processors - Modules or classes that process data or file representations.
                    These should contain methods for processing a model object, E.g. renumbering a file.
        > Presenters - Modules that bridge between a view and a process
                    These should drive a process from a form. E.g. A process that requires an input file and an output path.
        > Views - A form
                    These should only contain form logic, e.g. code that controls basic form controls.

Credits:
    Grant Haynes