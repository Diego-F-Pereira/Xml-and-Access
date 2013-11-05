'-------------------------------------------------------------------------------------------
' Name:         clsSaxErrorHandler
' Purpose:      Handles the errors of a SAX implementation
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-28-2013
' References:   Requires the Microsoft XML, v6.0 library
'-------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IVBSAXErrorHandler
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXErrorHandler_ignorableWarning(ByVal oLocator As MSXML2.IVBSAXLocator, _
                                               strErrorMessage As String, _
                                              ByVal nErrorCode As Long)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXErrorHandler_error(ByVal oLocator As MSXML2.IVBSAXLocator, _
                                    strErrorMessage As String, _
                                   ByVal nErrorCode As Long)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXErrorHandler_fatalError(ByVal oLocator As MSXML2.IVBSAXLocator, _
                                         strErrorMessage As String, _
                                        ByVal nErrorCode As Long)
'
End Sub
