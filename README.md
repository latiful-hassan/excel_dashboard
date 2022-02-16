# excel_dashboard
- A simple yet effective dasboard created with Excel with its goal being to summarise key customer order data with rudimentary analysis and being fully automated
- A drop-down menu was created with 'Validation' and upon selecting a customer, their details, order history and PivotTable is then updated and shown
- A PivotTable and customised Slicer is also shown for customer with availiable data in order to see their yearly sales
- VBA code was used in order to automate this as shown below:

  Public Sub UpdateOrderHistory()

      On Error GoTo noCustomerData

      ActiveWorkbook.SlicerCaches("Slicer_Order_Month").ClearManualFilter
      ActiveSheet.Shapes.Range(Array("Order Month")).Visible = True
      ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = True

      Dim pt As PivotTable        ' store the pivottable
      Dim field As PivotField     ' store reference to the pt filter field
      Dim newCustomer As String   ' store selected customer name

      Set pt = Worksheets("Yearly Orders PT").PivotTables("YearlyOrdersPT")
      Set field = pt.PivotFields("Customer Name")
      newCustomer = Worksheets("Customer Dashboard").Range("B3").Value

      With pt

          field.ClearAllFilters
          field.CurrentPage = newCustomer
          .RefreshTable

      End With

  pDone:
      Exit Sub

  noCustomerData:
      MsgBox ("Customer Does Not Exist!")

      ActiveSheet.Shapes.Range(Array("Order Month")).Visible = False
      ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = False

  End Sub

