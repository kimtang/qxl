
namespace qxl

module qxl =
    open System.Collections.Generic
    open ExcelDna.Integration
    open kx

    let rec stringify (r:kx.KObject) = 
        let abool x =
            match x with
            | true -> "1"
            | false -> "0"
        match r with
        | kx.Bool(x) -> match x with
                        | true -> "1b"
                        | false -> "0b"
        | kx.Byte(x) -> x |> string 
        | kx.Guid(x) -> x |> string
        | kx.Short(x) -> x |> string
        | kx.Int(x) -> x |> string
        | kx.Long(x) -> x |> string
        | kx.Real(x) -> x |> string
        | kx.Float(x) -> x |> string
        | kx.Char(x) -> x |> string
        | kx.String(x) -> "`"+x 
        | kx.Date(x) -> x.ToString("yyyy.MM.dd")
        | kx.Timestamp(x) -> x.ToString("yyyy.MM.ddDHH:mm:ss.ffffzzz")
        | kx.Minute(x) -> x.ToString("HH:mm")
        | kx.Second(x) -> x.ToString("HH:mm:ss")
        | kx.Month(x) -> x.ToString("yyyy.MM")
        | kx.DateTime(x) -> x.ToString("yyyy.MM.ddDHH:mm:ss.fff")
        | kx.TimeSpan(x) -> x.ToString("HH:mm:ss.fff")
        | kx.KTimespan(x) -> x.ToString("HH:mm:ss.ffffzzz")
        | kx.AKObject(x) ->
            match x.Length with
            | 0 -> "()"
            | 1 -> "enlist " + stringify x.[0]
            | _ -> x |> List.map stringify |> List.reduce(fun x y -> x  + ";" + y) |> (fun x -> "(" + x + ")" ) 
        | kx.ABool(x) ->
            match x.Length with
            | 0 -> "`bool$()"
            | 1 -> "enlist " + (kx.Bool(x.[0]) |> stringify )
            | _ -> x  |> Array.map abool |> Array.reduce(fun x y -> x + y) |> (fun x -> x + "b" )        
        | kx.AByte(x) ->
            match x.Length with
            | 0 -> "`byte$()"
            | 1 -> "enlist " + (kx.Byte(x.[0]) |> stringify ) 
            | _ -> x  |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)

        | kx.AGuid(x) ->
            match x.Length with
            | 0 -> "`guid$()"
            | 1 -> "enlist " + (kx.Guid(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + ";" + y) |> (fun x -> "(" + x + ")" )

        | kx.AShort(x) ->
            match x.Length with
            | 0 -> "`short()"
            | 1 -> "enlist " + (kx.Short(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AInt(x) ->
            match x.Length with
            | 0 -> "`int()"
            | 1 -> "enlist " + (kx.Int(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
        | kx.ALong(x) ->
            match x.Length with
            | 0 -> "`long()"
            | 1 -> "enlist " + (kx.Long(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AReal(x) ->
            match x.Length with
            | 0 -> "`real()"
            | 1 -> "enlist " + (kx.Real(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AFloat(x) ->
            match x.Length with
            | 0 -> "`float()"
            | 1 -> "enlist " + (kx.Float(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AChar(x) ->
            match x.Length with
            | 0 -> "`char()"
            | 1 -> "enlist " + (kx.Char(x.[0]) |> stringify )
            | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  +  y)|> (fun x -> "\"" + x + "\"" )
        | kx.AString(x) ->
            match x.Length with
            | 0 -> "`$()"
            | 1 -> "enlist " + (kx.String(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x-> "`"+x) |> Array.reduce(fun x y -> x  + y) 
        | kx.ADate(x) ->
            match x.Length with
            | 0 -> "`date()"
            | 1 -> "enlist " + (kx.Date(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.dd")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.ATimestamp(x) ->
            match x.Length with
            | 0 -> "`timestamp()"
            | 1 -> "enlist " + (kx.Timestamp(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.ddDHH:mm:ss.ffffzzz")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AMinute(x) ->
            match x.Length with
            | 0 -> "`minute()"
            | 1 -> "enlist " + (kx.Minute(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("HH:mm")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.ASecond(x) ->
            match x.Length with
            | 0 -> "`second()"
            | 1 -> "enlist " + (kx.Second(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("HH:mm:ss")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AMonth(x) ->
            match x.Length with
            | 0 -> "`month()"
            | 1 -> "enlist " + (kx.Month(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM")) |> Array.reduce(fun x y -> x  + " " + y) 
        | kx.ADateTime(x) ->
            match x.Length with
            | 0 -> "`datetime()"
            | 1 -> "enlist " + (kx.DateTime(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.ddDHH:mm:ss.fff")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.AKTimespan(x) ->
            match x.Length with
            | 0 -> "`timespan()"
            | 1 -> "enlist " + (kx.KTimespan(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("HH:mm:ss.fff")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.ATimeSpan(x) ->
            match x.Length with
            | 0 -> "`timespan()"
            | 1 -> "enlist " + (kx.TimeSpan(x.[0]) |> stringify )
            | _ -> x |> Array.map (fun x -> x.ToString("HH:mm:ss.ffffzzz")) |> Array.reduce(fun x y -> x  + " " + y)
        | kx.Dict(x,y) ->
            let sx = stringify x
            let sy = stringify y
            sx + "!" + sy

        | kx.Flip(x,y) -> 
            let sx = 
                match x.Length with
                | 0 -> "()"
                | 1 -> "enlist`"+ x.[0]
                | _ -> x |> Array.map(fun x -> "`"+x) |> Array.reduce(fun x y -> x + y)
            let sy = 
                match y.Length with
                |0 -> ""
                |1 -> "enlist " + stringify y.[0]
                |_ -> y |> List.map stringify |> List.reduce(fun x y -> x  + ";" + y) |> (fun x -> "(" + x + ")" )

            "flip "+sx + "!" + sy
        | _ -> "nyi" 

    let rec conv_as_array (k:kx.KObject) =
        match k with
        | ABool(x) -> x |> Array.map (fun x -> x :> obj)
        | AGuid(x) -> x |> Array.map (fun x -> x |> string :> obj)
        | AByte(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | AShort(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | AInt(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | ALong(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | AReal(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | AFloat(x) -> x |> Array.map (fun x -> x |> float :> obj)
        | AChar(x) -> x |> Array.map (fun x -> x |> string :> obj)
        | AString(x) -> x |> Array.map (fun x -> x :> obj)
        | ATimestamp(x) -> x |> Array.map (fun x -> x :> obj)
        | AMonth(x) -> x |> Array.map (fun x -> x :> obj)
        | ADate(x) -> x |> Array.map (fun x -> x :> obj)
        | ADateTime(x) -> x |> Array.map (fun x -> x :> obj)
        | AKTimespan(x) -> x |> Array.map (fun x -> x :> obj)
        | AMinute(x) -> x |> Array.map (fun x -> x :> obj)
        | ASecond(x) -> x |> Array.map (fun x -> x :> obj)
        | ATimeSpan(x) -> x |> Array.map (fun x -> x :> obj)
        | Flip(x,y) -> 
            let xs = x |> Array.map (fun x -> x :> obj)
            let yc = y |> List.map (conv_as_array ) |> List.toArray
            // get the size an merge it together
            [|1 :> obj|] 



        | Dict(x,y) -> 
            // let xc = conv_as_matrix  x
            // let yc = conv_as_matrix  y
            // get the size and merge it together
            [|1 :> obj|] 
        | AKObject(x) -> x |> List.map ( fun x -> 
                                            match x with
                                            | Bool(x) -> x :> obj
                                            | Guid(x) -> x |> string :> obj
                                            | Byte(x) -> x |> float :> obj
                                            | Short(x) -> x |> float :> obj 
                                            | Int(x) -> x |> float :> obj
                                            | Long(x) -> x |> float :> obj
                                            | Real(x) -> x |> float :> obj
                                            | Float(x) -> x :> obj
                                            | Char(x) -> x :> obj
                                            | String(x) -> x :> obj 
                                            | Timestamp(x) -> x :> obj
                                            | Month(x) -> x :> obj
                                            | Date(x) -> x :> obj
                                            | DateTime(x) -> x :> obj
                                            | KTimespan(x) -> x :> obj
                                            | Minute(x) -> x :> obj
                                            | Second(x) -> x :> obj
                                            | TimeSpan(x) -> x :> obj
                                            | _ -> stringify x :> obj 
                                        ) |> List.toArray 
        | _ -> [|1 :> obj|] 


    // this is now used in table and dic
    let rec conv_as_matrix (k:kx.KObject) =
        match k with
        | ABool(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | AGuid(x) -> x |> Array.map (fun x -> x |> string :> obj) |> Array.create 1
        | AByte(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | AShort(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | AInt(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | ALong(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | AReal(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | AFloat(x) -> x |> Array.map (fun x -> x |> float :> obj) |> Array.create 1
        | AChar(x) -> x |> Array.map (fun x -> x |> string :> obj) |> Array.create 1
        | AString(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | ATimestamp(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | AMonth(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | ADate(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | ADateTime(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | AKTimespan(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | AMinute(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | ASecond(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | ATimeSpan(x) -> x |> Array.map (fun x -> x :> obj) |> Array.create 1
        | Flip(x,y) -> 
            let xs = x |> Array.map (fun x -> [|x :> obj|]) 
            let yc = y |> List.map (conv_as_array ) |> List.toArray
            // get the size an merge it together
            let r = yc |> Array.zip xs |> Array.map (fun (x,y) -> x |> Array.append y)
            r



        | Dict(x,y) -> 
            let xc = conv_as_matrix  x
            let yc = conv_as_matrix  y
            // get the size and merge it together
            [|1 :> obj|] |> Array.create 1

        | AKObject(x) -> let x = x |> List.map conv_as_array
                         let mn = x |> List.map (fun x -> x.Length) |> List.max
                         let r = x |> List.map (fun x -> [|for i in 1 .. (mn - x.Length) -> "" :> obj |] |> Array.append x) |> List.toArray
                         r

        | _ -> [|1 :> obj|] |> Array.create 1 

    let ktox (k:kx.KObject) = 
        match k with
        | Bool(x) -> x :> obj
        | Guid(x) -> x |> string :> obj
        | Byte(x) -> x |> float :> obj
        | Short(x) -> x |> float :> obj 
        | Int(x) -> x |> float :> obj
        | Long(x) -> x |> float :> obj
        | Real(x) -> x |> float :> obj
        | Float(x) -> x :> obj
        | Char(x) -> x :> obj
        | String(x) -> "`"+x :> obj 
        | Timestamp(x) -> x :> obj
        | Month(x) -> x :> obj
        | Date(x) -> x :> obj
        | DateTime(x) -> x :> obj
        | KTimespan(x) -> x :> obj
        | Minute(x) -> x :> obj
        | Second(x) -> x :> obj
        | TimeSpan(x) -> x :> obj
        | ABool(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | AGuid(x) -> x |> Array.map (fun x -> x |> string :> obj) :> obj
        | AByte(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | AShort(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | AInt(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | ALong(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | AReal(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | AFloat(x) -> x |> Array.map (fun x -> x |> float :> obj) :> obj
        | AChar(x) -> x |> Array.map (fun x -> x |> string :> obj) :> obj
        | AString(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | ATimestamp(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | AMonth(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | ADate(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | ADateTime(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | AKTimespan(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | AMinute(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | ASecond(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | ATimeSpan(x) -> x |> Array.map (fun x -> x :> obj) :> obj
        | Flip(x,y) -> 
            let xs = x |> Array.map (fun x -> [|x :> obj|]) 
            let yc = y |> List.map (conv_as_array ) |> List.toArray
            // get the size an merge it together
            let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
            
            let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
            
            twoDimensionalArray :> obj

        | Dict(x,y) -> 
            let xc = conv_as_matrix x
            let yc = conv_as_matrix y

            let yc = match yc.Length with
                     | 1 -> yc
                     | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
            // let yc2 = Array2D.init r.[0].Length r.Length (fun i j -> r.[j].[i])

            // let r = yc |> Array.zip xc |> Array.map (fun (x,y) -> y |> Array.append x)

            let r = yc |> Array.append xc 

            // let m = r |> Array.map (fun x -> x.Length) |> Array.max
            let twoDimensionalArray = Array2D.init r.[0].Length r.Length   (fun i j -> r.[j].[i])
            twoDimensionalArray :> obj

        | AKObject(x) -> let x = x |> List.map conv_as_array
                         let mn = x |> List.map (fun x -> x.Length) |> List.max
                         let r = x |> List.map (fun x -> [|for i in 1 .. (mn - x.Length) -> "" :> obj |] |> Array.append x) |> List.toArray
                         let twoDimensionalArray = Array2D.init r.Length r.[0].Length  (fun i j -> r.[i].[j])
                         twoDimensionalArray :> obj
        | _ -> 1 :> obj

    type XObject =
        | ExcelMissing
        | KObject of kx.KObject

    let xtok (arg:obj) = 
        match arg with
        | :? ExcelEmpty -> ExcelMissing
        | :? ExcelError -> ExcelMissing
        | :? ExcelMissing -> ExcelMissing
        | :? bool as o-> KObject(kx.KObject.Bool(o))
        | :? string as o-> KObject(kx.KObject.String(o))
        | :? float as o-> KObject(kx.KObject.Float(o))
        | :? (obj[,]) as o->
                let r = o |> Array2D.map (fun y ->
                                            match y with 
                                            | :? ExcelEmpty -> kx.KObject.String("ExcelEmpty")
                                            | :? ExcelError -> kx.KObject.String("ExcelError")
                                            | :? ExcelMissing -> kx.KObject.String("ExcelMissing")
                                            | :? bool as o-> kx.KObject.Bool(o)
                                            | :? string as o-> kx.KObject.String(o)
                                            | :? float as o-> kx.KObject.Float(o)
                                            )
                let s = [ for i in  0 .. r.GetLength(0) - 1 -> r.[i,*] ] |> List.map (fun x -> kx.KObject.AKObject(x |> Array.toList))
                KObject(kx.KObject.AKObject(s))
        | _ -> KObject(kx.KObject.Bool(true))

    let connectionMaps =
        new Dictionary<string,kx.c>()

    [<ExcelFunction(Description="My first .NET function")>]
    let dnaDescribe (arg:obj) =
        match arg with
        | :? System.DateTime as o-> "DateTime"
        | :? ExcelEmpty as o-> "ExcelEmpty"
        | :? ExcelError as o-> "ExcelError"
        | :? ExcelMissing as o-> "ExcelMissing"
        | :? bool as o-> "bool"
        | :? string as o-> "string"
        | :? float as o-> "float"
        | :? (obj[]) as o-> "1d array"
        | :? (obj[,]) as o-> "2d array:" + string( o.GetLength(0)) + "," + string(o.GetLength(1))
        | _ -> "no match"


    [<ExcelFunction(Description="open connection")>]
    let open_connection (uid:string) (host:string) (port:int) (user:string) =
        let con = match connectionMaps.ContainsKey uid with
                  | false -> 
                        let con = kx.c(host,port,user)
                        connectionMaps.Add(uid,con)
                        con
                  | true -> connectionMaps.Item uid
        
        uid

    [<ExcelFunction(Description="execute query")>]
    let execute (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) (c:obj) = 

        match connectionMaps.ContainsKey uid with
        | false -> "uid not found" :> obj
        | true -> let con = connectionMaps.Item uid
                  let x,y,z,a,b,c = xtok x,xtok y,xtok z,xtok a,xtok b,xtok c
                  
                  let r = match x,y,z,a,b,c with
                          | ExcelMissing,_,_,_,_,_-> con.k(query)
                          | KObject(x),ExcelMissing,_,_,_,_-> con.k(query,x)
                          | KObject(x),KObject(y),ExcelMissing,_,_,_-> con.k(query,x,y)
                          | KObject(x),KObject(y),KObject(z),ExcelMissing,_,_-> con.k(query,x,y,z)
                          | KObject(x),KObject(y),KObject(z),KObject(a),ExcelMissing,_-> con.k(query,x,y,z,a)
                          | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),ExcelMissing-> con.k(query,x,y,z,a,b)
                          | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),KObject(c)-> con.k(query,x,y,z,a,b,c)
                  ktox r

    [<ExcelFunction(Description="execute query")>]
    let close_connection (uid:string)  =
        match connectionMaps.ContainsKey uid with
        | false -> ()
        | true -> connectionMaps.Item(uid).close()

        uid |> connectionMaps.Remove 
