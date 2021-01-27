REM  *****  LIBREOFFICE WRITER v6.4.6.2 BASIC  *****
REM Macro to paste an image as character.
REM You can assign it to a keyboard shortcut, for example CTRL-SHIFT-V



sub Main
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SetAnchorToChar", "", 0, Array())


end sub
