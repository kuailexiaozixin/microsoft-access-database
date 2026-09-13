Version =20
VersionRequired =20
Begin Form
    FilterOn = NotDefault
    DefaultView =0
    ViewsAllowed =1
    Width =10700
    Filter ="ID = 1"
    RecSrcDt = Begin
        0xdeee4806b67de640
    End
    RecordSource ="tOrders"
    Caption ="Order Detail"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            BorderColor =12632256
        End
        Begin Section
            Height =5400
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =200
                    Width =1500
                    Height =315
                    Name ="lblID"
                    Caption ="ID"
                    LayoutCachedLeft =200
                    LayoutCachedTop =200
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =515
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1900
                    Top =200
                    Width =2000
                    Height =315
                    Name ="txtID"
                    ControlSource ="ID"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =200
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =515
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =650
                    Width =1500
                    Height =315
                    Name ="lblCustomerID"
                    Caption ="Customer"
                    LayoutCachedLeft =200
                    LayoutCachedTop =650
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =965
                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =3200
                    Left =1900
                    Top =650
                    Width =4000
                    Height =315
                    TabIndex =1
                    Name ="cboCustomerID"
                    ControlSource ="CustomerID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tCustomers.ID, tCustomers.FirstName & \" \" & tCustomers.LastName AS Cust"
                        "omerName FROM tCustomers ORDER BY tCustomers.LastName, tCustomers.FirstName; "
                    ColumnWidths ="0;3000"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =650
                    LayoutCachedWidth =5900
                    LayoutCachedHeight =965
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =1100
                    Width =1500
                    Height =315
                    Name ="lblOrderNumber"
                    Caption ="Order Number"
                    LayoutCachedLeft =200
                    LayoutCachedTop =1100
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =1415
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =1100
                    Width =3000
                    Height =315
                    TabIndex =2
                    Name ="txtOrderNumber"
                    ControlSource ="OrderNumber"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =1100
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =1415
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =1550
                    Width =1500
                    Height =315
                    Name ="lblOrderDate"
                    Caption ="Order Date"
                    LayoutCachedLeft =200
                    LayoutCachedTop =1550
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =1865
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =1550
                    Width =3000
                    Height =315
                    TabIndex =3
                    Name ="txtOrderDate"
                    ControlSource ="OrderDate"
                    Format ="Short Date"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =1550
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =1865
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =2100
                    Width =1500
                    Height =315
                    Name ="lblOrderItems"
                    Caption ="Order Items"
                    LayoutCachedLeft =200
                    LayoutCachedTop =2100
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =2415
                End
                Begin Subform
                    OverlapFlags =85
                    Top =2500
                    Width =10700
                    Height =2800
                    TabIndex =4
                    Name ="subOrderItems"
                    SourceObject ="Form.fOrderItemsSub"
                    LinkChildFields ="OrderID"
                    LinkMasterFields ="ID"

                    LayoutCachedTop =2500
                    LayoutCachedWidth =10700
                    LayoutCachedHeight =5300
                End
            End
        End
        Begin FormFooter
            Height =600
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =200
                    Top =100
                    Width =2500
                    Height =440
                    Name ="cmdSaveClose"
                    Caption ="Save and Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =200
                    LayoutCachedTop =100
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2900
                    Top =100
                    Width =2500
                    Height =440
                    TabIndex =1
                    Name ="cmdDiscardClose"
                    Caption ="Discard and Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2900
                    LayoutCachedTop =100
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5600
                    Top =100
                    Width =2500
                    Height =440
                    TabIndex =2
                    Name ="cmdPreview"
                    Caption ="Preview Order"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5600
                    LayoutCachedTop =100
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =540
                End
            End
        End
    End
End
CodeBehindForm
' See "fOrderDetail.cls"
