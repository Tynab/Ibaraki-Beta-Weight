﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Ibaraki_Beta_Weight.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Ibaraki Beta Weight.
        '''</summary>
        Friend ReadOnly Property app_name() As String
            Get
                Return ResourceManager.GetString("app_name", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to 茨城 (ベタ) 重量.
        '''</summary>
        Friend ReadOnly Property app_true_name() As String
            Get
                Return ResourceManager.GetString("app_true_name", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to 22.10.10.
        '''</summary>
        Friend ReadOnly Property app_ver() As String
            Get
                Return ResourceManager.GetString("app_ver", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to COPYRIGHT by Yami An.
        '''</summary>
        Friend ReadOnly Property cc_text() As String
            Get
                Return ResourceManager.GetString("cc_text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Nephilim.
        '''</summary>
        Friend ReadOnly Property co_name() As String
            Get
                Return ResourceManager.GetString("co_name", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to エマールグループ.
        '''</summary>
        Friend ReadOnly Property gr_name() As String
            Get
                Return ResourceManager.GetString("gr_name", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap.
        '''</summary>
        Friend ReadOnly Property gUpdating() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("gUpdating", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to WUFOIOiMqOWfjiAo44OZ44K/KSDph43ph48=.
        '''</summary>
        Friend ReadOnly Property key_ser() As String
            Get
                Return ResourceManager.GetString("key_ser", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to https://raw.githubusercontent.com/Tynab/Tynab/main/app/Ibaraki%20Beta%20Weight.
        '''</summary>
        Friend ReadOnly Property link_app() As String
            Get
                Return ResourceManager.GetString("link_app", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to https://www.google.com/.
        '''</summary>
        Friend ReadOnly Property link_base() As String
            Get
                Return ResourceManager.GetString("link_base", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to https://raw.githubusercontent.com/Tynab/Tynab/main/ver/Ibaraki%20Beta%20Weight.
        '''</summary>
        Friend ReadOnly Property link_ver() As String
            Get
                Return ResourceManager.GetString("link_ver", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
