' VBAPorter --- provide a interface of import/export VBA source file on Excel

' Copyright (C) 2014  Hiroaki Otsu

' Author: Hiroaki Otsu <ootsuhiroaki@gmail.com>
' URL: https://github.com/aki2o/vba-porter
' Version: 0.0.1

' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' Enjoy!!!


Option Explicit

Private Sub Workbook_AddinInstall()
    VBAPorter.initialize True
End Sub

Private Sub Workbook_AddinUninstall()
    VBAPorter.finalize True
End Sub

Private Sub Workbook_Open()
    VBAPorter.initialize True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Cancel Then Exit Sub
    VBAPorter.finalize True
End Sub


