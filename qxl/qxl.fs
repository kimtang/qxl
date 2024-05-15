
namespace qxl

open System
open System.Threading
open ExcelDna.Integration
open ExcelDna.Registration.FSharp

module qxl =
    open System.Collections.Generic
    open kx
    open ktox

    type XObject =
        | ExcelMissing
        | KObject of kx.KObject

    type XType = | XExcelEmpty | XExcelError | XExcelMissing | XBool | XString | XFloat | XMix

    type XParse = |Star|B|G|X|H|I|J|E|F|C|S|P|M|D|Z|N|U|V|T

    let getXParse0 (x:string) =
        let s = x.Split(".")
        if 1=s.Length then x,Star else
        let l = s |> Array.rev |> Array.head |> fun x -> x.ToUpper()
        let x1 = s |> Array.rev |> Array.tail |> Array.rev |> Array.reduce (fun x y -> x + "." + y)
        match l with
        | "*" -> x1,Star
        | "B" -> x1,B
        | "G" -> x1,I
        | "X" -> x1,I
        | "H" -> x1,I
        | "I" -> x1,I
        | "J" -> x1,J
        | "E" -> x1,E
        | "F" -> x1,F
        | "C" -> x1,C
        | "S" -> x1,S
        | "P" -> x1,P
        | "M" -> x1,M
        | "D" -> x1,D
        | "Z" -> x1,Z
        | "N" -> x1,N
        | "U" -> x1,U
        | "V" -> x1,V
        | "T" -> x1,T
        | _ -> x,Star


    let getXParse (x:string array) =  
        let x0 = x |> Array.map getXParse0
        let x1 = x0 |>  Array.map fst
        let x2 = x0 |>  Array.map snd
        x1,x2

    let xtok1dim (arg: obj array) =
        let cnt = Array.length arg
        let r= arg |> Array.map ( fun y -> 
                                            match y with
                                            | :? ExcelEmpty -> XExcelEmpty 
                                            | :? ExcelError -> XExcelError 
                                            | :? ExcelMissing -> XExcelMissing
                                            | :? bool as o-> XBool
                                            | :? string as o-> XString
                                            | :? float as o-> XFloat
                                            ) |> Array.distinct |> Array.length
        match r with
        | 1 ->  let oo = arg[0]
                let r = match oo with
                        | :? ExcelEmpty -> XExcelEmpty,Array.create cnt "ExcelEmpty" |> kx.KObject.AString
                        | :? ExcelError -> XExcelError,Array.create cnt "ExcelError" |> kx.KObject.AString
                        | :? ExcelMissing -> XExcelMissing,Array.create cnt "ExcelMissing" |> kx.KObject.AString 
                        | :? bool -> XBool ,arg |> Array.map (fun x -> x :?> bool) |> kx.KObject.ABool
                        | :? string -> XString ,arg |> Array.map (fun x -> x :?> string) |> kx.KObject.AString
                        | :? float -> XFloat ,arg |> Array.map (fun x -> x :?> float) |> kx.KObject.AFloat
                r
        | _ -> let r = arg |> Array.map ( fun y -> 
                                        match y with
                                        | :? ExcelEmpty -> kx.KObject.String("ExcelEmpty")
                                        | :? ExcelError -> kx.KObject.String("ExcelError")
                                        | :? ExcelMissing -> kx.KObject.String("ExcelMissing")
                                        | :? bool as o-> kx.KObject.Bool(o)
                                        | :? string as o-> kx.KObject.String(o)
                                        | :? float as o-> kx.KObject.Float(o)
                                        ) |> Array.toList |> kx.KObject.AKObject
               XMix,r        

    let xtok1dimStar (arg:obj array) =
        let r = match arg.Length with
                | 1 -> match arg.[0] with
                       | :? ExcelEmpty -> XExcelEmpty,kx.KObject.String("ExcelEmpty")
                       | :? ExcelError -> XExcelError,kx.KObject.String("ExcelError")
                       | :? ExcelMissing -> XExcelMissing,kx.KObject.String("ExcelMissing")
                       | :? bool as o-> XBool ,kx.KObject.Bool(o)
                       | :? string as o-> XString ,kx.KObject.String(o)
                       | :? float as o-> XFloat,kx.KObject.Float(o)
                | _ -> xtok1dim arg
        r

    let xtok1dimB (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> o
                                           | :? string as o-> match o with
                                                              | "1b" -> true
                                                              | "0b" -> false
                                                              | "TRUE" -> true
                                                              | "False" -> false
                                           | :? float as o-> if o<1 then false else true
                                           | _ -> false
                         )
        match r.Length with
        | 1 -> XBool,kx.KObject.Bool(r.[0])
        | _ -> XBool,kx.KObject.ABool(r)


    let xtok1dimG (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.Guid.Empty
                                           | :? string as o-> 
                                                    try
                                                        System.Guid.Parse(o)
                                                    with
                                                        | _ -> System.Guid.Empty

                                           | :? float as o-> System.Guid.Empty
                                           | _ -> System.Guid.Empty
                         )
        match r.Length with
        | 1 -> XBool,kx.KObject.Guid(r.[0])
        | _ -> XFloat,kx.KObject.AGuid(r)
        

    let xtok1dimX0 (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1uy else 0uy
                                           | :? string as o-> 
                                                    try
                                                        byte(o)
                                                    with
                                                        | _ -> 0uy

                                           | :? float as o-> byte(o)
                                           | _ -> 0uy
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Byte(r.[0])
        | _ -> XFloat,kx.KObject.AByte(r)

    let xtok1dimH (arg:obj array) =
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1s else 0s
                                           | :? string as o-> 
                                                    try
                                                        int16(o)
                                                    with
                                                        | _ -> 0s

                                           | :? float as o-> int16(o)
                                           | _ -> 0s
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Short(r.[0])
        | _ -> XFloat,kx.KObject.AShort(r)    
        
    let xtok1dimI (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1 else 0
                                           | :? string as o-> 
                                                    try
                                                        int(o)
                                                    with
                                                        | _ -> 0

                                           | :? float as o-> int(o)
                                           | _ -> 0
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Int(r.[0])
        | _ ->XFloat,kx.KObject.AInt(r)

    let xtok1dimJ (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1L else 0L
                                           | :? string as o-> 
                                                    try
                                                        int64(o)
                                                    with
                                                        | _ -> 0L

                                           | :? float as o-> int64(o)
                                           | _ -> 0L
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Long(r.[0])
        | _ -> XFloat,kx.KObject.ALong(r)

    let xtok1dimE (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1.0f else 0.0f
                                           | :? string as o-> 
                                                    try
                                                        single(o)
                                                    with
                                                        | _ -> 0.0f

                                           | :? float as o-> single(o)
                                           | _ -> 0.0f
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Real(r.[0])
        | _ ->XFloat,kx.KObject.AReal(r)

    let xtok1dimF (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then 1.0 else 0.0
                                           | :? string as o-> 
                                                    try
                                                        float(o)
                                                    with
                                                        | _ -> 0.0

                                           | :? float as o-> float(o)
                                           | _ -> 0.0
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Float(r.[0])
        | _ ->XFloat,kx.KObject.AFloat(r)

    let xtok1dimC (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then '1' else '0'
                                           | :? string as o-> 
                                                    match o.Length with
                                                    | 0 -> '0'
                                                    | _ -> o.Chars(0)
                                                    
                                           | :? float as o-> '0'
                                           | _ -> '0'
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Char(r.[0])
        | _ ->XFloat,kx.KObject.AChar(r)

    let xtok1dimS (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> if o then "true" else "false"
                                           | :? string as o-> o
                                           | :? float as o-> string(o)
                                           | _ -> ""
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.String(r.[0])
        | _ ->XFloat,kx.KObject.AString(r)

    let xtok1dimP (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.DateTime(538589095631452241L)
                                           | :? string as o-> 
                                                    try
                                                        System.DateTime.Parse(o)
                                                    with
                                                        | _ -> System.DateTime(538589095631452241L)

                                           | :? float as o-> System.DateTime.FromOADate(o)
                                           | _ -> System.DateTime(538589095631452241L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Timestamp(r.[0])
        | _ ->XFloat,kx.KObject.ATimestamp(r)

    let xtok1dimM (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.DateTime(538589095631452241L)
                                           | :? string as o-> 
                                                    try
                                                        System.DateTime.Parse(o)
                                                    with
                                                        | _ -> System.DateTime(538589095631452241L)

                                           | :? float as o-> System.DateTime.FromOADate(o)
                                           | _ -> System.DateTime(538589095631452241L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Month(r.[0])
        | _ ->XFloat,kx.KObject.AMonth(r)

    let xtok1dimD (arg:obj array) =
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.DateTime(538589095631452241L)
                                           | :? string as o-> 
                                                    try
                                                        System.DateTime.Parse(o)
                                                    with
                                                        | _ -> System.DateTime(538589095631452241L)

                                           | :? float as o-> System.DateTime.FromOADate(o)
                                           | _ -> System.DateTime(538589095631452241L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Date(r.[0])
        | _ ->XFloat,kx.KObject.ADate(r)

    let xtok1dimZ (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.DateTime(538589095631452241L)
                                           | :? string as o-> 
                                                    try
                                                        System.DateTime.Parse(o)
                                                    with
                                                        | _ -> System.DateTime(538589095631452241L)

                                           | :? float as o-> System.DateTime.FromOADate(o)
                                           | _ -> System.DateTime(538589095631452241L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.DateTime(r.[0])
        | _ ->XFloat,kx.KObject.ADateTime(r)

    let xtok1dimN (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.TimeSpan(-21474836480000L)
                                           | :? string as o-> 
                                                    try
                                                        System.TimeSpan.Parse(o)
                                                    with
                                                        | _ -> System.TimeSpan(-21474836480000L)

                                           | :? float as o-> System.TimeSpan.FromSeconds(o)
                                           | _ -> System.TimeSpan(-21474836480000L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.TimeSpan(r.[0])
        | _ ->XFloat,kx.KObject.ATimeSpan(r)

    let xtok1dimU (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.TimeSpan(-21474836480000L)
                                           | :? string as o-> 
                                                    try
                                                        System.TimeSpan.Parse(o)
                                                    with
                                                        | _ -> System.TimeSpan(-21474836480000L)

                                           | :? float as o-> System.TimeSpan.FromSeconds(o)
                                           | _ -> System.TimeSpan(-21474836480000L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Minute(r.[0])
        | _ ->XFloat,kx.KObject.AMinute(r)

    let xtok1dimV (arg:obj array) =
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.TimeSpan(-21474836480000L)
                                           | :? string as o-> 
                                                    try
                                                        System.TimeSpan.Parse(o)
                                                    with
                                                        | _ -> System.TimeSpan(-21474836480000L)

                                           | :? float as o-> System.TimeSpan.FromSeconds(o)
                                           | _ -> System.TimeSpan(-21474836480000L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.Second(r.[0])
        | _ ->XFloat,kx.KObject.ASecond(r)

    let xtok1dimT (arg:obj array) = 
        let r = arg |> Array.map (fun x -> match x with
                                           | :? bool as o-> System.TimeSpan(-21474836480000L)
                                           | :? string as o-> 
                                                    try
                                                        System.TimeSpan.Parse(o)
                                                    with
                                                        | _ -> System.TimeSpan(-21474836480000L)

                                           | :? float as o-> System.TimeSpan.FromSeconds(o)
                                           | _ -> System.TimeSpan(-21474836480000L)
                         )
        match r.Length with
        | 1 -> XFloat,kx.KObject.KTimespan(r.[0])
        | _ ->XFloat,kx.KObject.AKTimespan(r)

    let xtok1dimX (arg: obj array)(t:XParse) =
        match t with
        |Star -> xtok1dimStar arg
        |B    -> xtok1dimB arg
        |G    -> xtok1dimG arg
        |X    -> xtok1dimX0 arg
        |H    -> xtok1dimH arg
        |I    -> xtok1dimI arg
        |J    -> xtok1dimJ arg
        |E    -> xtok1dimE arg
        |F    -> xtok1dimF arg
        |C    -> xtok1dimC arg
        |S    -> xtok1dimS arg
        |P    -> xtok1dimP arg
        |M    -> xtok1dimM arg
        |D    -> xtok1dimD arg
        |Z    -> xtok1dimZ arg
        |N    -> xtok1dimN arg
        |U    -> xtok1dimU arg
        |V    -> xtok1dimV arg
        |T    -> xtok1dimT arg

    // we want to turn a matrix either in dic or tbl
    // let's check
    let xtok2dim (arg: obj[,]) =
        let r = arg[0,*] |> xtok1dim
        let c = arg[*,0] |> xtok1dim
        let rcnt = arg |> Array2D.length1
        let lcnt = arg |> Array2D.length2
        
        match fst r,fst c with
        | XString,_ -> 
                        let s,d = match (r  |> snd) with
                                  | kx.KObject.AString(s0) -> getXParse s0
                                  | _ -> (Array.create rcnt "ExcelEmpty"),Array.create rcnt Star
                        let a = d |> Array.zip [|0 .. lcnt - 1|]
                                |> Array.map (fun (i,d) -> xtok1dimX arg[1 .. ,i] d) |> Array.map snd
                                |> Array.toList
                        kx.Flip(s, a)
        | _,XString -> 

                        let s,d = match (c  |> snd) with
                                  | kx.KObject.AString(s0) -> getXParse s0
                                  | _ -> (Array.create rcnt "ExcelEmpty"),Array.create rcnt Star

                        let a = d |> Array.zip [|0 .. rcnt - 1|]
                                |> Array.map (fun (i,d) -> xtok1dimX arg[i,1 .. ] d) |> Array.map snd
                                |> Array.toList
                                |> kx.KObject.AKObject
                        
                        let k = s |> kx.KObject.AString

(*                        let a = [0 .. rcnt - 1] |> List.map ( fun i -> let ab = arg[i,1 .. ]
                                                                               |> Array.map ( fun x ->  match x with
                                                                                                        | :? ExcelEmpty -> false,x
                                                                                                        | :? ExcelError -> false,x
                                                                                                        | :? ExcelMissing -> false,x
                                                                                                        | _ -> true,x
                                                                                            )
                                                                               |> Array.filter (fun (b,x) -> b)
                                                                               |> Array.map (fun (b,x) -> x)
                                                                       match ab.Length with
                                                                       | 1 ->   match ab.[0] with
                                                                                | :? bool as o-> kx.KObject.Bool(o)
                                                                                | :? string as o-> kx.KObject.String(o)
                                                                                | :? float as o-> kx.KObject.Float(o)
                                                                                | _ ->  kx.KObject.Bool(true)

                                                                       | _ -> ab |> xtok1dim |> snd
                                                            ) 
                                                |> kx.KObject.AKObject*)
                        kx.Dict(k, a)
        | _,_ -> 
                        let r = arg |> Array2D.map (fun y ->
                                                        match y with 
                                                        | :? ExcelEmpty -> kx.KObject.String("ExcelEmpty")
                                                        | :? ExcelError -> kx.KObject.String("ExcelError")
                                                        | :? ExcelMissing -> kx.KObject.String("ExcelMissing")
                                                        | :? bool as o-> kx.KObject.Bool(o)
                                                        | :? string as o-> kx.KObject.String(o)
                                                        | :? float as o-> kx.KObject.Float(o)
                                                        )
                        let s = [ for i in  0 .. r.GetLength(0) - 1 -> r.[i,*] ] |> List.map (fun x -> kx.KObject.AKObject(x |> Array.toList))
                        kx.KObject.AKObject(s)

    let xtok (arg:obj) = 
        match arg with
        | :? ExcelEmpty -> ExcelMissing
        | :? ExcelError -> ExcelMissing
        | :? ExcelMissing -> ExcelMissing
        | :? bool as o-> KObject(kx.KObject.Bool(o))
        | :? string as o-> KObject(kx.KObject.String(o))
        | :? float as o-> KObject(kx.KObject.Float(o))
        | :? (obj[,]) as o->
                let rcnt = o |> Array2D.length1
                let lcnt = o |> Array2D.length2
                match rcnt,lcnt with
                | _,1  ->   o[*,0] |> xtok1dim |> snd |>  KObject
                | 1,_  ->   o[0,*] |> xtok1dim |> snd |> KObject
                | _,_  -> o |> xtok2dim  |> KObject

        | _ -> KObject(kx.KObject.Bool(true))

    let connectionMaps =
        new Dictionary<string,kx.c>()

    [<ExcelFunction(Description="My first .NET function")>]
    let dna_connection(arg:obj) =
        connectionMaps.Keys |> Seq.toArray :> obj


    [<ExcelFunction(Description="My first .NET function")>]
    let dna_cell_reference() =
        let reference:ExcelReference = XlCall.Excel(XlCall.xlfCaller) |> unbox
        let cellReference:string = XlCall.Excel(XlCall.xlfAddress, 1+reference.RowFirst,1+reference.ColumnFirst) |> unbox
        let sheetName:string = XlCall.Excel(XlCall.xlSheetNm,reference) |> unbox
        sheetName + cellReference

    [<ExcelFunction(Description="My first .NET function")>]
    let dna_describe (arg:obj) =
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
    let dna_open_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)=
        match port with
        | 0 -> ExcelMissing :> obj
        | _ ->

            let uid =   match connectionMaps.ContainsKey uid with
                        | false ->  
                                    try 
                                        let con = kx.c(host,port,user+":"+passwd)                                    
                                        connectionMaps.Add(uid,con)
                                        uid
                                    with
                                    | ex -> "no_con" 

                        | true ->   let con = connectionMaps.Item uid
                                    match con.Connected() with
                                    | true -> uid
                                    | false ->  connectionMaps.Remove uid |> ignore
                                                try
                                                    let con = kx.c(host,port,user+":"+passwd)
                                                    connectionMaps.Add(uid,con)
                                                    uid
                                                with
                                                | ex -> "no_con"
        
            uid :> obj
             
    let csMaps =
        new Dictionary<string,kx.cs>()

    let async_open_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)=
        async {
            match port with
            | 0 -> return ExcelMissing :> obj
            | _ ->
                    match csMaps.ContainsKey uid with
                    | false -> 
                                let con = new kx.cs()
                                let! msg = con.Init(host,port,user+":"+passwd)

                                match msg with
                                | "connected" -> csMaps.Add(uid,con) |> ignore
                                                 return uid :> obj
                                | _ -> return "no_con" :> obj                                          

                    | true ->   let con = connectionMaps.Item uid
                                match con.Connected() with
                                | true -> return  uid :> obj                                                 
                                | false ->  csMaps.Remove uid |> ignore
                                            let con = new kx.cs()
                                            let! msg = con.Init(host,port,user+":"+passwd)
                                            match msg with
                                            | "connected" -> csMaps.Add(uid,con) |> ignore
                                                             return uid :> obj
                                            | _ -> return "no_con" :> obj
                                            
        }

    let rec argComparer1 (x:obj) (y:obj) : bool =
        match x,y with
        | :? ExcelEmpty as x, (:? ExcelEmpty as y)-> true
        | :? ExcelError as x,(:? ExcelError as y)-> true
        | :? ExcelMissing as x,(:? ExcelMissing as y)-> true
        | :? bool as x,(:? bool as y)-> x=y
        | :? string as x,(:? string as y)-> x.Equals y
        | :? float as x,(:? float as y)-> x=y
        | :? (obj[]) as x,(:? (obj[]) as y)->
            let dx,dy = x.Length,y.Length
            if not (dx=dy) then false else argComparer2 x y
        | :? (obj[,]) as x,(:? (obj[,]) as y)->
            let dx1,dx2 = Array2D.length1 x,Array2D.length2 x
            let dy1,dy2 = Array2D.length1 y,Array2D.length2 y
            if not((dx1=dy1) && (dx2=dy2)) then false else argComparer3 x y
        | _,_ -> false
    and argComparer2 (x:obj[]) (y:obj[]) =
        x |> Array.zip y |> Array.map (fun (x,y) -> argComparer1 x y ) |> Array.reduce (&&)
    and argComparer3 (x:obj[,]) (y:obj[,]) =
        let s = [| for i in  0 .. x.GetLength(0) - 1 -> argComparer2 x.[i,*] y.[i,*] |] |> Array.reduce (&&)
        s
        

    let argComparer0 (x1,x2,x3,x4,x5,x6,x7) (y1,y2,y3,y4,y5,y6,y7) : bool =
        argComparer1 [|x1;x2;x3;x4;x5;x6;x7|] [|y1;y2;y3;y4;y5;y6;y7|]

    type argComparer() = 
        interface IEqualityComparer<obj*string*obj*obj*obj*obj*obj*obj> with
            member this.Equals((x1,x2,x3,x4,x5,x6,x7,x8),(y1,y2,y3,y4,y5,y6,y7,y8)): bool = 
                let z0 = argComparer0 (x1,x2,x3,x4,x5,x6,x7) (y1,y2,y3,y4,y5,y6,y7)
                z0 && x2.Equals y2
            member this.GetHashCode((x1,x2,x3,x4,x5,x6,x7,x8)): int = 
                System.HashCode.Combine(x1,x2,x3,x4,x5,x6,x7,x8)

        
    // fun (x1,x2,x3,x4,x5,x6,x7,x8) (y1,y2,y3,y4,y5,y6,y7,y8) -> true 

    let argMaps =
        new Dictionary<string,obj*string*obj*obj*obj*obj*obj*obj*obj*System.DateTime>()

    [<ExcelFunction(Description="async execute query")>]
    let dna_aopen_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)= 
        FsAsyncUtil.excelRunAsync "async_open_connection" [|uid :> obj; host :> obj; port :> obj; user :> obj ; passwd :> obj |] (async_open_connection uid host port user passwd )

    [<ExcelFunction(Description="execute query")>]
    let dna_execute (random:obj) (uid:string) (query:obj) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 

        let query = match query with
                    | :? string as query-> query
                    | :? (obj[,]) as o ->
                        let rcnt = o |> Array2D.length1
                        let lcnt = o |> Array2D.length2
                        let arr = match rcnt,lcnt with
                                  | _,1  ->   o[*,0] |> Array.map (fun x -> 
                                                                         match x with
                                                                         | :? string as s -> true,s
                                                                         | _ -> false,""
                                                                )
                                                     |> Array.filter (fun (x,s) -> x)
                                                     |> Array.map (fun (x,s) -> s)
                                                     |> Array.fold (fun x y -> x + "\n" + y) ""
                                  | 1,_  ->   o[0,*] |> Array.map (fun x -> 
                                                                         match x with
                                                                         | :? string as s -> true,s
                                                                         | _ -> false,""
                                                                )
                                                     |> Array.filter (fun (x,s) -> x)
                                                     |> Array.map (fun (x,s) -> s)
                                                     |> Array.fold (fun x y -> x + "\n" + y) ""

                                  | _,_  ->  let arr = [| for i in  0 .. o.GetLength(0) - 1 -> o.[i,*] |]
                                                       |> Array.map (
                                                            fun x -> x |> Array.map(    fun x0 -> match x0 with
                                                                                                  | :? string as s -> s
                                                                                                  | :? bool as b -> if b then "1b" else "0b"
                                                                                                  | :? float as f -> if f=System.Math.Round(f) then sprintf "%i" (int (System.Math.Round(f))) else sprintf "%f" f
                                                                                                                     
                                                                                                  | _ -> ""
                                                                                        ) |> Array.fold (fun x y -> x + y) "" 
                                                            )
                                             arr |> Array.fold (fun x y -> x + "\n" + y) ""
                        arr
                    | _ -> failwith "not a vector"

        match connectionMaps.ContainsKey uid with
        | false -> "uid not found" :> obj
        | true -> let con = connectionMaps.Item uid
                  let x1,y1,z1,a1,b1 = xtok x,xtok y,xtok z,xtok a,xtok b
                  
                  let r = match x1,y1,z1,a1,b1 with
                          | ExcelMissing,_,_,_,_-> con.k(query)
                          | KObject(x1),ExcelMissing,_,_,_-> con.k(query,x1)
                          | KObject(x1),KObject(y1),ExcelMissing,_,_-> con.k(query,x1,y1)
                          | KObject(x1),KObject(y1),KObject(z1),ExcelMissing,_-> con.k(query,x1,y1,z1)
                          | KObject(x1),KObject(y1),KObject(z1),KObject(a1),ExcelMissing-> con.k(query,x1,y1,z1,a1)
                          | KObject(x1),KObject(y1),KObject(z1),KObject(a1),KObject(b1)-> con.k(query,x1,y1,z1,a1,b1)
                          // | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),KObject(c)-> con.k(query,x,y,z,a,b,c)
                  let o = ktox r

                  let o1 = match o,x with
                           | :? (obj[]) as o0 , (:? (obj[,]) as x0)->
                                        let rcnt = x0 |> Array2D.length1
                                        let lcnt = x0 |> Array2D.length2
                                        match rcnt,lcnt with
                                        | 1,_  -> o
                                        | _,1  -> let tmp = o0 |> Array.map (fun x -> [|x|]) |> array2D
                                                  tmp :> obj
                                        | _,_  -> o
                           | _ -> o
                  o1

    let asyncExecute (asCols:bool) (arg:string) (random) (uid:string) (query:obj) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) =
        async {
                // let arg = random,uid,query,x,y,z,a,b
                let query = match query with
                            | :? string as query-> query
                            | :? (obj[,]) as o ->
                                let rcnt = o |> Array2D.length1
                                let lcnt = o |> Array2D.length2
                                let arr = match rcnt,lcnt with
                                          | _,1  ->   o[*,0] |> Array.map (fun x -> 
                                                                                 match x with
                                                                                 | :? string as s -> true,s
                                                                                 | _ -> false,""
                                                                        )
                                                             |> Array.filter (fun (x,s) -> x)
                                                             |> Array.map (fun (x,s) -> s)
                                                             |> Array.fold (fun x y -> x + "\n" + y) ""
                                          | 1,_  ->   o[0,*] |> Array.map (fun x -> 
                                                                                 match x with
                                                                                 | :? string as s -> true,s
                                                                                 | _ -> false,""
                                                                        )
                                                             |> Array.filter (fun (x,s) -> x)
                                                             |> Array.map (fun (x,s) -> s)
                                                             |> Array.fold (fun x y -> x + "\n" + y) ""

                                          | _,_  ->  let arr = [| for i in  0 .. o.GetLength(0) - 1 -> o.[i,*] |]
                                                               |> Array.map (
                                                                    fun x -> x |> Array.map(    fun x0 -> match x0 with
                                                                                                          | :? string as s -> s
                                                                                                          | :? bool as b -> if b then "1b" else "0b"
                                                                                                          | :? float as f -> if f=System.Math.Round(f) then sprintf "%i" (int (System.Math.Round(f))) else sprintf "%f" f
                                                                                                                     
                                                                                                          | _ -> ""
                                                                                                ) |> Array.fold (fun x y -> x + y) "" 
                                                                    )
                                                     arr |> Array.fold (fun x y -> x + "\n" + y) ""
                                arr
                            | _ -> failwith "not a vector"

                match csMaps.ContainsKey uid with
                | false -> return "uid not found" :> obj
                | true -> let con = csMaps.Item uid
                          let x1,y1,z1,a1,b1 = xtok x,xtok y,xtok z,xtok a,xtok b                  
                          let! r =match x1,y1,z1,a1,b1 with
                                  | ExcelMissing,_,_,_,_-> con.k(query)
                                  | KObject(x1),ExcelMissing,_,_,_-> con.k(query,x1)
                                  | KObject(x1),KObject(y1),ExcelMissing,_,_-> con.k(query,x1,y1)
                                  | KObject(x1),KObject(y1),KObject(z1),ExcelMissing,_-> con.k(query,x1,y1,z1)
                                  | KObject(x1),KObject(y1),KObject(z1),KObject(a1),ExcelMissing-> con.k(query,x1,y1,z1,a1)
                                  | KObject(x1),KObject(y1),KObject(z1),KObject(a1),KObject(b1)-> con.k(query,x1,y1,z1,a1,b1)
                                  // | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),KObject(c)-> con.k(query,x,y,z,a,b,c)
                          let o = ktox r

                          let o1 = match o,x with
                                   | :? (obj[]) as o0 , (:? (obj[,]) as x0)->
                                        let rcnt = x0 |> Array2D.length1
                                        let lcnt = x0 |> Array2D.length2
                                        match rcnt,lcnt with
                                        | 1,_  -> o
                                        | _,1  -> let tmp = o0 |> Array.map (fun x -> [|x|]) |> array2D
                                                  tmp :> obj
                                        | _,_  -> o
                                   | _ -> o

                          let o2 = match o1,asCols with
                                   | :? (obj[]) as o0, true -> let tmp = o0 |> Array.map (fun x -> [|x|]) |> array2D
                                                               tmp :> obj
                                   | :? (obj[,]) as o0,false -> o1
                                   | _ -> o1

                          if argMaps.ContainsKey arg then arg |> argMaps.Remove |> ignore
                          argMaps.Add(arg,(random,uid,query,x,y,z,a,b,o2,System.DateTime.Now))
                          return o2
        }

    [<ExcelFunction(Description="async execute query")>]
    let dna_aexecute (random:obj) (uid:string) (query:obj) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) =
        let arg = dna_cell_reference()
        // let arg = random,uid,query,x,y,z,a,b
        match argMaps.ContainsKey arg with
        | false -> FsAsyncUtil.excelRunAsync "asyncExecute" [|(false :> obj);(arg :> obj); random; uid; query; x; y; z; a; b |] (asyncExecute false arg random uid query x y z a b )
        | true -> let (random1,uid1,query1,x1,y1,z1,a1,b1,r,p) = argMaps.Item arg
                  if uid.Equals uid1 && argComparer0 (random,query,x,y,z,a,b) (random1,query1,x1,y1,z1,a1,b1) then r 
                  else 
                    FsAsyncUtil.excelRunAsync "asyncExecute" [|(false :> obj);(arg :> obj); random; uid; query; x; y; z; a; b |] (asyncExecute false arg random uid query x y z a b )
                  // let t = System.DateTime.Now - p
                  // if t.TotalSeconds < 1 then o else FsAsyncUtil.excelRunAsync "asyncExecute" [|random; (uid :> obj); (query); x; y; z; a; b |] (asyncExecute random uid query x y z a b )

    [<ExcelFunction(Description="async execute query as column")>]
    let dna_aexecutec (random:obj) (uid:string) (query:obj) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) =
        let arg = dna_cell_reference()
        // let arg = random,uid,query,x,y,z,a,b
        match argMaps.ContainsKey arg with
        | false -> FsAsyncUtil.excelRunAsync "asyncExecute" [|(true :> obj);(arg :> obj); random; uid; query; x; y; z; a; b |] (asyncExecute true arg random uid query x y z a b )
        | true -> let (random1,uid1,query1,x1,y1,z1,a1,b1,r,p) = argMaps.Item arg
                  if uid.Equals uid1 && argComparer0 (random,query,x,y,z,a,b) (random1,query1,x1,y1,z1,a1,b1) then r 
                  else 
                    FsAsyncUtil.excelRunAsync "asyncExecute" [|(true :> obj);(arg :> obj); random; uid; query; x; y; z; a; b |] (asyncExecute true arg random uid query x y z a b )
                  // let t = System.DateTime.Now - p
                  // if t.TotalSeconds < 1 then o else FsAsyncUtil.excelRunAsync "asyncExecute" [|random; (uid :> obj); (query); x; y; z; a; b |] (asyncExecute random uid query x y z a b )

        

    [<ExcelFunction(Description="execute query")>]
    let dna_close_connection (uid:string)  =
        match connectionMaps.ContainsKey uid with
        | false -> ()
        | true -> connectionMaps.Item(uid).close()

        uid |> connectionMaps.Remove 
                    
    // Helper that will create a timer that ticks at timerInterval for timerDuration, then stops
    // Not exported to Excel (incompatible type)
    let createTimer timerInterval timerDuration =
        // setup a timer
        let startTime = DateTime.Now
        let timer = new System.Timers.Timer(float timerInterval)
        timer.AutoReset <- true
        // return an async task for stopping
        let timerStop = async {
            timer.Start()
            do! Async.Sleep (timerDuration |> int)
            timer.Stop() 
            }
        Async.Start timerStop
        // Make sure that the type we observe in the event is supported by Excel
        // (events like timer.Elapsed are automatically IObservable in F#)
        timer.Elapsed |> Observable.map (fun elapsed -> [|"S: "+startTime.ToString("HH:mm:ss.fff");"G: "+DateTime.Now.ToString("HH:mm:ss.fff")|] |> Array.map(fun x -> x :> obj) :> obj) 

    // Excel function to start the timer - using the fact that F# events implement IObservable
    [<ExcelFunction(Description="start timer")>]
    let dna_startTimer timerInterval timerDuration =
        FsAsyncUtil.excelObserve "startTimer" [|float timerInterval; float timerDuration|] (createTimer timerInterval timerDuration)


    [<ExcelFunction(Description="Create subscriber")>]
    let dna_open_subscriber (uid:string) (host:string) (port:int) (user:string) (passwd:string) (sub:string) (append:int)=
        match port with
        | 0 -> ExcelMissing :> obj
        | _ ->
            match rtd.RtdKdb.subscriberMaps.ContainsKey uid with
            | false ->
                  let con = rtd.RtdKdb.kdb_subscriber(uid,host,port,user+":"+passwd,sub,append)
                  rtd.RtdKdb.subscriberMaps.Add(uid,con)               
            | true ->
                  rtd.RtdKdb.subscriberMaps.[uid].Close()
                  rtd.RtdKdb.subscriberMaps.[uid] <- rtd.RtdKdb.kdb_subscriber(uid,host,port,user+":"+passwd,sub,append)
            uid :> obj
        


    [<ExcelFunction(Description="subscribe to table and sym")>]
    let dna_subscribe (uid:string) =
        let r = XlCall.RTD("rtdkdb.rtdkdbserver","",uid)
        match rtd.RtdKdb.objMaps.ContainsKey uid with
        | true -> rtd.RtdKdb.objMaps.[uid]
        | false -> r 

    [<ExcelFunction(Description="Provides rtd access")>]
    let dna_rtd(progId:string) (server:string) (topic:string)=
        XlCall.RTD(progId,server,topic)

    [<ExcelFunction(Description="Provides a ticking clock")>]
    let dna_clock (progId:obj) =
        XlCall.RTD("rtdclock.rtdclockserver","","rtdclock.rtdclockserver")

