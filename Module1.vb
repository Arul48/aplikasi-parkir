Imports System.Data.Odbc
Imports System.Data
Imports System.Data.OleDb
Module Module1
    Public OLECMD As OleDbCommand
    Public OLERDR As OleDbDataReader
    Public OLEDA As OleDbDataAdapter
    Public CNN As OleDbConnection
    Public DS As DataSet
    Public KONEKSI As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\DB Zery 1.accdb"
    Public x As Integer

End Module

