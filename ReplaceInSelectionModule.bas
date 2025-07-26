Attribute VB_Name = "ReplaceInSelectionModule"
Option Explicit

' =====================================================================
'   2025-07-26 - Initial Version
'
'   Outlook VBA Macro to Find and Replace text in the current Selection of an Outlook Item.
'
'   For help on using the VBA Editor, Self Signing and running Macros, setting a reference to the Word Object library,
'   and adding Macros to your Quick Access Toolbar or Ribbon - See the Slipstick Systems web site article:
'   How to use Outlook's VBA Editor  https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/
'
'   You must add a Reference to the "Microsoft Word X.xx Object Library" to compile.
'
'   Copyright (C) 2024, 2025 Lycon Of Texas
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
'   License Version 3 as published by the Free Software Foundation. This program is distributed in the hope that it
'   will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
'   PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the
'   GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.
'
' =====================================================================

'   Main Sub (Macro)
'
Public Sub ReplaceInSelection()

    Word_ReplaceInSelection "Find This", "Replace With This"
    
End Sub

'   Do a Word Find/Replace on the current Selection
'
Private Function Word_ReplaceInSelection(ByVal FindText As String, ByVal ReplaceText As String) As Boolean
Word_ReplaceInSelection = False

    '   Must have an Active Inspector
    '
    If Not (TypeOf ActiveWindow Is Outlook.Inspector) Then
        MsgBox "The Active Window must be an Inspector.", vbExclamation
        Exit Function
    End If
    
    '   Get the Active Inspector Word Doc
    '
    Dim wDoc As Word.Document
    Set wDoc = ActiveInspector.WordEditor
            
    '   Word Doc must be editable
    '
    If wDoc Is Nothing Then
        MsgBox "Active Inspector has no Word Editor.", vbExclamation
        Exit Function
    End If
    
    If wDoc.ProtectionType <> wdNoProtection Then
        MsgBox "Active Inspector is Locked For Editing (Read Only).", vbExclamation
        Exit Function
    End If
    
    '   Word Doc must have a Selection
    '
    Dim wDocSelection As Word.Selection
    Set wDocSelection = wDoc.Application.Selection
    If wDocSelection Is Nothing Then
        MsgBox "Active Inspector Selection is Nothing.", vbExclamation
        Exit Function
    End If
    If wDocSelection.Start = wDocSelection.End Then
        MsgBox "Active Inspector Selection is empty.", vbExclamation
        Exit Function
    End If
    
    '   Replace all occurances
    '
    Dim wDocSearch As Word.Range
    Set wDocSearch = Word_FindDefault(wDocSelection.Range.Duplicate)
    wDocSearch.Find.Text = FindText
    wDocSearch.Find.Replacement.Text = ReplaceText
    wDocSearch.Find.Execute Replace:=wdReplaceAll

Word_ReplaceInSelection = True
End Function

'   Reset a Word .Find object to defaults
'
'   From https://gregmaxey.com/word_tip_pages/words_fickle_vba_find_property.html
'
Private Function Word_FindDefault(ByVal wRange As Word.Range) As Word.Range

    Set Word_FindDefault = wRange
    With Word_FindDefault.Find
    
        .ClearFormatting
        .Format = False
        .Forward = True
        .Highlight = wdUndefined
        .IgnorePunct = False
        .IgnoreSpace = False
        .MatchAllWordForms = False
        .MatchCase = False
        .MatchPhrase = False
        .MatchPrefix = False
        .MatchSoundsLike = False
        .MatchSuffix = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Text = ""
        .Wrap = wdFindStop

    End With
    
End Function

