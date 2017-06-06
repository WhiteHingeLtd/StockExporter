﻿Imports WHLClasses

Module Module1

    Sub Main()

        Console.WriteLine("Stock Replenishment CSV Exporter.")
        Console.WriteLine("This program is good for the enivronment - It was made with 94% Recycled code.")
        LocationsWithStockProdutProxy()
        LocationsWithStockVariantProxy()
        Console.WriteLine("Jobs done. Have a nice day!")
        Threading.Thread.Sleep(10000)
    End Sub

    Private Sub SaveCSV(Data As  List(Of Dictionary(Of String, Object)), Filename as String )
        dim rawcsv as String = ""

        rawcsv += String.Join(",", Data(0).Keys)
        My.Computer.FileSystem.WriteAllText(Filename, rawcsv, False)
        rawcsv = ""
        Console.WriteLine("Writing data")
        Dim i As Integer = 0
        For Each row As Dictionary(Of String, Object) In Data
            'T("RowStart")
            i += 1
            'sender.ReportProgress((i / Data.Count) * 100)
            'T("Prog")
            rawcsv += vbNewLine


            If rawcsv.Length > 10240 Then
                rawcsv = rawcsv.Replace("\", "\\").Replace("""", """""").Replace("," + vbNewLine, vbNewLine).Replace("§", """")
                try
                    My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
                    rawcsv = ""
                Catch ex As Exception
                    Console.WriteLine("File in use, delaying save")
                End Try
            End If
            
            'T("LineStart")
            For Each value In row.Values
                If IsNothing(value)
                    rawcsv += ","
                Else 
                    If value.ToString.Contains(",") Or value.ToString.Contains("""") Or value.ToString.Contains(vbNewLine) Or value.ToString.StartsWith(" ") Or value.ToString.EndsWith(" ") Then
                    rawcsv += "§" + value.ToString.Replace(vbNewLine, "") + "§,"
                Else
                    rawcsv += value.ToString.Replace(vbNewLine, "") + ","
                End If
                End If
                
            Next
            'T("Linedone")
        Next
        rawcsv = rawcsv.Replace("\", "\\").Replace("""", """""").Replace("," + vbNewLine, vbNewLine).Replace("§", """")
        'sender.ReportProgress(100, "Saving file")
        try
            My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Catch ex As Exception
            Console.WriteLine("File in use, delaying save")
            Threading.Thread.Sleep("20000")
            try
            My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Catch ex2 As Exception
                Console.WriteLine("File in use, delaying save")
            Threading.Thread.Sleep("20000")
                try
            My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Catch ex3 As Exception
                    Console.WriteLine("File in use, delaying save")
            Threading.Thread.Sleep("20000")
                    try
            My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Catch ex4 As Exception
                        Console.WriteLine("File in use, delaying save")
            Threading.Thread.Sleep("20000")
                        try
            My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Catch ex5 As Exception
                            Console.WriteLine("File in use, Giving Up")
        End Try
        End Try
        End Try
        End Try
        End Try
        My.Computer.FileSystem.WriteAllText(Filename, rawcsv, True)
        Console.WriteLine("""" + Filename + """ Written.")
    End Sub

    Private Sub LocationsWithStockProdutProxy()
        Console.WriteLine("=== PRODUCTS ===")
        Console.WriteLine("Starting Server Side Processing.")
        Dim Iterates As Integer = MySQL.SelectDataDictionary("SELECT Count(*) as count FROM whldata.sku_locations Group BY Sku ORDER BY Count(*) DESC LIMIT 1;")(0)("count")

        Dim Skus As List(Of Dictionary(Of String, Object)) = MySQL.SelectDataDictionary("
SELECT	
	a.Sku,
    a.ItemTitle,
    c.whltotal as 'Stock_Total',  
    d.Shelfname as 'PickLocation',
    d.additionalinfo as 'PickStockLevel',
    (e.stock + e.stockminimum) as 'LinnworksStock',
    (c.whltotal-(e.stock + e.stockminimum)) as 'StockDiff'
FROM whldata.whlnew as a
LEFT JOIN (SELECT Sku, Sum(additionalinfo*CAST(substring(sku,8,4) as signed integer)) as 'whltotal' from whldata.sku_locations group by substring(sku,1,7)) as c on a.Sku=c.sku
LEFT JOIN (SELECT Sku, Shelfname, Sum(additionalinfo*CAST(substring(sku,8,4) as signed integer)) as additionalinfo FROM whldata.sku_locations JOIN whldata.locationreference on sku_locations.LocationRefID=locationreference.LocID WHERE LocType=1 Group by substring(sku,1,7)) as d on a.Sku=d.Sku
LEFT JOIN whldata.inventory as e on a.Sku=e.Sku
WHERE 	(NOT New_Status='Dead') AND  (IsListed='True' OR Packsize=1) AND (NOT a.IsBundle='True') AND (HasBeenListed='True' or New_Status='Exported') AND  ( a.sku LIKE '%0001');

")

        Dim Locations As List(Of Dictionary(Of String, Object)) = MySQL.SelectDataDictionary("SELECT shelfname, Sku, substring(sku,1,7) as ShortSku, SUM(additionalInfo*CAST(SUBSTRING(Sku,8,4) as signed integer)) as additionalinfo, locationRefID, if(locType=0,99,locType) as 'type',locWarehouse as Warehouse FROM whldata.sku_locations as a JOIN whldata.locationReference as b on b.locId=a.LocationRefId WHERE NOT locType=1 group by shortsku, shelfname;")

        Dim Fields As New List(Of String)
        Fields.Add("sku")
        Fields.Add("ItemTitle")
        Fields.Add("StockTotal")
        Fields.Add("PickLocation")
        Fields.Add("PickStock")
        Fields.Add("LinnworksTotal")
        Fields.Add("Difference")
        For I As Integer = 1 To Iterates
            Fields.Add("Shelf_" + I.ToString)
            'Fields.Add("Stocklevel_" + i.tostring)
        Next

        Dim data As New List(Of Dictionary(Of String, Object))

        'Now we can iterate through and sort them out
        Dim IterCount As Integer = 0
        Console.WriteLine("Starting Client side processing.")
        For Each Sku As Dictionary(Of String, Object) In Skus
            IterCount += 1
            'Worker.ReportProgress((IterCount / Skus.Count) * 100, "Loading ""Locations on Skus"" Data... (" + IterCount.ToString + " of " + Skus.Count.ToString + ")")


            Dim NewRow As New Dictionary(Of String, Object)
            'Create the fields
            For Each Field As String In Fields
                NewRow.Add(Field, Nothing)
            Next
            'Now fill them. Start witht he easy one.
            NewRow("sku") = Sku("Sku")
            NewRow("StockTotal") = Sku("Stock_Total")
            NewRow("ItemTitle") = Sku("ItemTitle")
            NewRow("PickLocation") = Sku("PickLocation")
            NewRow("PickStock") = Sku("PickStockLevel")
            NewRow("LinnworksTotal") = Sku("LinnworksStock")
            NewRow("Difference") = Sku("StockDiff")
            'Gte the locations which apply
            Dim RelevantLocations As List(Of Dictionary(Of String, Object)) = Locations.Where(Function(x As Dictionary(Of String, Object)) x("Sku") = Sku("Sku")).ToList
            RelevantLocations.Sort(Function(x As Dictionary(Of String, Object), y As Dictionary(Of String, Object)) x("type").CompareTo(y("type")))
            Dim LocationNumber As Integer = 0
            For Each Location As Dictionary(Of String, Object) In RelevantLocations
                LocationNumber += 1
                NewRow("Shelf_" + LocationNumber.ToString) = Location("shelfname")
                'NewRow("Stocklevel_" + locationNumber.tostring) = Location("additionalInfo")
            Next
            data.Add(NewRow)
        Next
        'And now we can feed it in!
        Console.WriteLine("Saving Data File.")
        SaveCSV(data, "\\server\Data Storage\Shared Data\Reporting\Replenishment_Products.csv")
    End Sub

    Private Sub LocationsWithStockVariantProxy()
        Console.WriteLine("=== VARIANTS ===")
        Console.WriteLine("Starting Server Side Processing.")
                dim Iterates As Integer = Mysql.SelectDataDictionary("SELECT Count(*) as count FROM whldata.sku_locations Group BY Sku ORDER BY Count(*) DESC LIMIT 1;")(0)("count")
        
        Dim Skus as List(Of Dictionary(Of String, Object)) = MySQL.SelectDataDictionary("
SELECT	
	a.Sku,
    a.ItemTitle,
    c.whltotal as 'Stock_Total',  
    d.Shelfname as 'PickLocation',
    d.additionalinfo as 'PickStockLevel',
    e.ow_isprepackfinalfinal as 'Packdown'
FROM whldata.whlnew as a
LEFT JOIN (SELECT Sku, Sum(additionalinfo) as 'whltotal' from whldata.sku_locations group by sku) as c on a.Sku=c.sku
LEFT JOIN (SELECT Sku, Shelfname, additionalinfo FROM whldata.sku_locations JOIN whldata.locationreference on sku_locations.LocationRefID=locationreference.LocID WHERE LocType=1 Group by sku) as d on a.Sku=d.Sku
LEFT JOIN whldata.orderwise_data as e on a.Sku=e.Sku
WHERE 	(NOT New_Status='Dead') AND  (IsListed='True' OR Packsize=1) AND (NOT a.IsBundle='True') AND (HasBeenListed='True' or New_Status='Exported') AND  (Not a.sku LIKE '%xxxx');

")
      
        Dim Locations As List(Of Dictionary(Of String, Object)) = MySQL.SelectDataDictionary("SELECT shelfname, Sku, additionalInfo, locationRefID, if(locType=0,99,locType) as 'type',locWarehouse as Warehouse FROM whldata.sku_locations as a JOIN whldata.locationReference as b on b.locId=a.LocationRefId WHERE NOT locType=1;")
        
        dim Fields As new list(of String)
        Fields.Add("sku")
        Fields.Add("ItemTitle")
        Fields.Add("StockTotal")
        Fields.Add("PickLocation")
        Fields.Add("PickStock")
        Fields.Add("Packdown")
        For I as Integer = 1 to iterates
            Fields.Add("Shelf_" + i.tostring)
            Fields.Add("Stocklevel_" + i.tostring)
        Next

        Dim data as New List(Of Dictionary(Of String, Object))

        'Now we can iterate through and sort them out
        dim IterCount as Integer =0 
        Console.WriteLine("Starting Client side processing.")
        For each Sku as Dictionary(Of String, Object) in Skus
            IterCount += 1
            'Worker.ReportProgress((IterCount/skus.Count)*100, "Loading ""Locations on Skus"" Data... (" + Itercount.tostring + " of " +skus.Count.tostring+ ")" )
            
                                                                              
            Dim NewRow as New Dictionary(Of String, Object)
            'Create the fields
            For each Field as String in Fields
                NewRow.Add(Field,nothing)
            Next
            'Now fill them. Start witht he easy one.
            NewRow("sku") = sku("Sku")
            NewRow("StockTotal") = sku("Stock_Total")
            NewRow("ItemTitle") = sku("ItemTitle")
            NewRow("PickLocation") = sku("PickLocation")
            NewRow("PickStock") = Sku("PickStockLevel")
            NewRow("Packdown") = Sku("Packdown")
            'Gte the locations which apply
            Dim RelevantLocations As List(Of Dictionary(Of String, Object)) = Locations.Where(Function(x as Dictionary(Of String, Object)) x("Sku")=Sku("Sku")).ToList
            RelevantLocations.Sort(Function(x As Dictionary(Of String, Object), y As Dictionary(Of String, Object)) x("type").CompareTo(y("type")))
            dim LocationNumber as Integer = 0
            For each Location as Dictionary(Of String, Object) in RelevantLocations
                locationNumber += 1
                NewRow("Shelf_" + locationNumber.tostring) = Location("shelfname")
                NewRow("Stocklevel_" + locationNumber.tostring) = Location("additionalInfo")
            Next
            data.Add(newrow)
        Next
        'And now we can feed it in!
        Console.WriteLine("Saving Data File.")
        SaveCSV(data, "\\server\Data Storage\Shared Data\Reporting\Replenishment_Variants.csv")
    End Sub

End Module