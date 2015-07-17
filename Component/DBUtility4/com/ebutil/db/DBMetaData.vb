Option Explicit On
Option Strict On

Public Interface DBMetaData
    Function getMetaArray() As System.Collections.ArrayList
    Function getPrimaryKeys() As System.Collections.ArrayList
    Sub setPrimaryKeys()
    Sub init()
End Interface
