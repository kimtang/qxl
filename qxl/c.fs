
module kx
open System.Net.Sockets
open System

exception KException of string

// 2014.09.07 DONE: TODO: Add Dict and Flip Support
// 2014.09.07 DONE: TODO: Add Time support
// 2014.09.08 DONE: TODO: Add Exception Support
// 2014.09.08 DONE: TODO: Add maxbuffer size support
// 2014.09.08 DONE: Add Guid support
// 2020.04.09 DONE: Fixed month nan 
// TODO: extend qn function

// 06.09.2014: Position Bug fixed

let e = System.Text.Encoding.ASCII
let tt() = DateTime.Now.TimeOfDay
let nt = [0; 1; 16; 0; 1; 2; 4; 8; 4; 8; 1; 0; 8; 4; 4; 8; 8; 4; 4; 4]
let ni,nj,nf = Int32.MinValue,Int64.MinValue,Double.NaN
let o = 8.64e11 * 730119.0 |> int64
let odate = (new DateTime(2000,01,01,0,0,0,0)).ToOADate()
let ng = new Guid()
let za, zw = DateTime.MinValue.AddTicks(1L),DateTime.MaxValue
let clampDT j =  Math.Min(Math.Max(j, za.Ticks), zw.Ticks)

let ktofDate f =
        match f with
        | -2147483647 -> za
        | Int32.MaxValue -> zw
        | Int32.MinValue -> new DateTime(0L)
        | i -> let r = (i  + 730119 |> float) *8.64e11 |> int64
//               let r = clampDT((int64) (8.64e11 * ((i |> int64) + o |> float))) 
               new DateTime(r )

let ftoKDate (d:DateTime) =
        match d.Ticks with
        | 0L -> ni
        | x -> let r = x /(8.64e11 |> int64) - 730119L 
               r |> int

let ftoKMonth(x:System.DateTime) = (x.Year - 2000) * 12 + x.Month - 1

type KObject =
    | ERROR of string
    | TODO
    | NULL
    | Bool of bool
    | Guid of System.Guid
    | Byte of byte
    | Short of int16
    | Int of int
    | Long of int64
    | Real of single
    | Float of double
    | Char of char
    | String of string
    | Timestamp of System.DateTime
    | Month of System.DateTime
    | Date of System.DateTime
    | DateTime of System.DateTime
    | KTimespan of System.TimeSpan
    | Minute of System.TimeSpan
    | Second of System.TimeSpan
    | TimeSpan of System.TimeSpan
    | ABool of bool array
    | AGuid of System.Guid array
    | AByte of byte array
    | AShort of int16 array
    | AInt of int array
    | ALong of int64 array
    | AReal of single array
    | AFloat of double array
    | AChar of char array
    | AString of string array
    | ATimestamp of System.DateTime array
    | AMonth of System.DateTime array
    | ADate of System.DateTime array
    | ADateTime of System.DateTime array
    | AKTimespan of System.TimeSpan array
    | AMinute of System.TimeSpan array
    | ASecond of System.TimeSpan array
    | ATimeSpan of System.TimeSpan array
    | Dict of KObject * KObject
    | Flip of string array * KObject list
    | AKObject of KObject list

let at d s = 
    match d with
    | Flip(x,y) -> let i = Array.FindIndex(x,(fun x -> x.Equals s)) 
                   y.[i]
    | _ -> raise(KException("Error at function at"))

let td d = 
    match d with
    | Flip(x,y) -> d
    | Dict(x,y) -> match x,y with
                   | AString(a),AKObject(b) -> Flip(a,b)
                   | _,_ -> raise(KException("Flipping a non dictionary is not allowed"))
    | _ -> raise(KException("Flipping a non dictionary is not allowed"))



let qn k =
    match k with 
    | Bool(x) -> false |> x.Equals
    | Guid(x) -> ng|> x.Equals
    | Byte(x) -> 0uy |> x.Equals
    | Short(x) -> Int16.MinValue |>  x.Equals
    | Int(x) -> Int32.MinValue |> x.Equals
    | Long(x) -> Int64.MinValue |> x.Equals 
    | Real(x) -> nf |> float32 |> x.Equals
    | Float(x) -> nf |> x.Equals
    | Char(x) -> ' ' |> x.Equals
    | String(x) -> "" |> x.Equals
    | DateTime(x) -> new System.DateTime(0L) |> x.Equals
    | Month(x) -> ni = ftoKMonth x
    | Date(x) -> ni = ftoKDate x 
    | Timestamp(x) -> new System.DateTime(0L) |> x.Equals 
    | KTimespan(x) ->  100L*x.Ticks = nj
    | Minute(x) -> x.Hours*60 + x.Minutes = ni
    | Second(x) -> x.Hours*60*60 + x.Minutes*60 + x.Seconds = ni
    | TimeSpan(x) -> new System.TimeSpan(0L) |> x.Equals
    | _ -> false
    

let t(k:KObject) = 
    match k with
    | Bool(_) -> -1 
    | Guid(_) -> -2 
    | Byte(_) -> -4 
    | Short(_) -> -5 
    | Int(_) -> -6 
    | Long(_) -> -7 
    | Real(_) -> -8 
    | Float(_) -> -9 
    | Char(_) -> -10 
    | String(_) -> -11
    | Timestamp(_) -> -12
    | Month(_) -> -13
    | Date(_) -> -14
    | DateTime(_) -> -15
    | KTimespan(_) -> -16
    | Minute(_) -> -17
    | Second(_) -> -18
    | TimeSpan(_) -> -19
    | ABool(_) -> 1
    | AGuid(_) -> 2
    | AByte(_) -> 4
    | AShort(_) -> 5
    | AInt(_) -> 6
    | ALong(_) -> 7
    | AReal(_) -> 8
    | AFloat(_) -> 9
    | AChar(_) -> 10
    | AString(_) -> 0
    | ATimestamp(_) -> 12 
    | AMonth(_) ->  13
    | ADate(_) ->  14
    | ADateTime(_) -> 15  
    | AKTimespan(_) -> 16
    | AMinute(_) ->  17
    | ASecond(_) ->  18
    | ATimeSpan(_) ->  19
    | Flip(_,_) -> 98
    | Dict(_,_) -> 99
    | AKObject (_) -> 0
    | ERROR(s) -> raise (KException(s))
    | TODO -> raise (KException("TODO in t"))
    | NULL -> raise (KException("NULL in t"))

let rec n(k:KObject) = 
    match k with
    | Bool(_) -> 1 
    | Guid(_) -> 1 
    | Byte(_) -> 1 
    | Short(_) -> 1 
    | Int(_) -> 1 
    | Long(_) -> 1 
    | Real(_) -> 1 
    | Float(_) -> 1 
    | Char(_) -> 1 
    | String(_) -> 1
    | Timestamp(_) -> 1
    | Month(_) -> 1
    | Date(_) -> 1
    | DateTime(_) -> 1
    | KTimespan(_) -> 1
    | Minute(_) -> 1
    | Second(_) -> 1
    | TimeSpan(_) -> 1
    | ABool(x) -> Array.length x
    | AGuid(x) -> Array.length x
    | AByte(x) -> Array.length x
    | AShort(x) -> Array.length x
    | AInt(x) -> Array.length x
    | ALong(x) -> Array.length x
    | AReal(x) -> Array.length x
    | AFloat(x) -> Array.length x
    | AChar(x) -> Array.length x
    | AString(x) -> Array.length x
    | ATimestamp(x) -> Array.length x 
    | AMonth(x) -> Array.length x
    | ADate(x) -> Array.length x
    | ADateTime(x) -> Array.length x  
    | AKTimespan(x) -> Array.length x
    | AMinute(x) -> Array.length x
    | ASecond(x) -> Array.length x
    | ATimeSpan(x) -> Array.length x
    | Flip(x,y) -> n y.[0]
    | Dict(x,y) -> n x
    | AKObject (x) -> List.length x
    | ERROR(x) -> raise (KException(x))
    | TODO -> raise (KException("TODO in n"))
    | NULL -> raise (KException("NULL in n"))

let ns(s:string) = 
    let i = s.IndexOf('\000')
    let i = if -1 < i then i else s.Length
    s.Substring(0,i) |> e.GetBytes |> fun x -> x.Length

let rec nx k = 
    match k with
    | Bool(_) -> 1+1
    | Guid(_) -> 1+16
    | Byte(_) -> 1+1
    | Short(_) -> 1+2
    | Int(_) -> 1+4
    | Long(_) -> 1+8
    | Real(_) -> 1+4
    | Float(_) -> 1+8
    | Char(_) -> 1+1
    | String(s) -> 2 + ns s
    | Timestamp(_) -> 1+8
    | Month(_) -> 1+4
    | Date(_) -> 1+4
    | DateTime(_) -> 1+8
    | KTimespan(_) -> 1+8
    | Minute(_) -> 1+4
    | Second(_) -> 1+4
    | TimeSpan(_) -> 1+4
    | ABool(x) -> 6 + 1 * x.Length
    | AGuid(x) -> 6 + 16 * x.Length 
    | AByte(x) -> 6 + 1 * x.Length 
    | AShort(x) -> 6 + 2 * x.Length 
    | AInt(x) -> 6 + 4 * x.Length 
    | ALong(x) -> 6 + 8 * x.Length 
    | AReal(x) -> 6 + 4 * x.Length 
    | AFloat(x) -> 6 + 8 * x.Length 
    | AChar(x) -> 6 + 1 * x.Length 
    | AString(x) -> x |> Array.fold (fun s x-> s+2+ns x) 6
    | ATimestamp(x) -> 6 + 8 * x.Length 
    | AMonth(x) -> 6 + 4 * x.Length
    | ADate(x) -> 6 + 4 * x.Length
    | ADateTime(x) -> 6 + 8 * x.Length  
    | AKTimespan(x) -> 6 + 8 * x.Length
    | AMinute(x) -> 6 + 4 * x.Length
    | ASecond(x) -> 6 + 4 * x.Length
    | ATimeSpan(x) -> 6 + 4 * x.Length
    | Flip(x,y) -> 3 + nx(AString(x)) + nx(AKObject(y))
    | Dict(x,y) -> 1 + nx(x) + nx(y)
    | AKObject(x) -> x |> List.map nx |> List.fold (+) 6
    | ERROR(x) -> raise(KException(x))
    | NULL -> raise(KException("unknown error"))
    | TODO -> raise(KException("unknown todo"))

let rec stringify (r:KObject) = 
    let abool x =
        match x with
        | true -> "1"
        | false -> "0"
    match r with
    | ERROR(x) -> "'" + x
    | Bool(x) -> match x with
                    | true -> "1b"
                    | false -> "0b"
    | Byte(x) -> x |> string 
    | Guid(x) -> x |> string
    | Short(x) -> x |> string
    | Int(x) -> x |> string
    | Long(x) -> x |> string
    | Real(x) -> x |> string
    | Float(x) -> x |> string
    | Char(x) -> x |> string
    | String(x) -> "`"+x 
    | Date(x) -> x.ToString("yyyy.MM.dd")
    | Timestamp(x) -> x.ToString("yyyy.MM.ddDHH:mm:ss.ffffzzz")
    | Minute(x) -> x.ToString("hh\:mm")
    | Second(x) -> x.ToString("hh\:mm\:ss")
    | Month(x) -> x.ToString("yyyy.MM")
    | DateTime(x) -> x.ToString("yyyy.MM.ddDHH:mm:ss.fff")
    | TimeSpan(x) -> x.ToString("hh\:mm\:ss\.fff")
    | KTimespan(x) -> x.ToString("hh\:mm\:ss\.fffffff")
    | AKObject(x) ->
        match x.Length with
        | 0 -> "()"
        | 1 -> "enlist " + stringify x.[0]
        | _ -> x |> List.map stringify |> List.reduce(fun x y -> x  + ";" + y) |> (fun x -> "(" + x + ")" ) 
    | ABool(x) ->
        match x.Length with
        | 0 -> "`bool$()"
        | 1 -> "enlist " + (Bool(x.[0]) |> stringify )
        | _ -> x  |> Array.map abool |> Array.reduce(fun x y -> x + y) |> (fun x -> x + "b" )        
    | AByte(x) ->
        match x.Length with
        | 0 -> "`byte$()"
        | 1 -> "enlist " + (Byte(x.[0]) |> stringify ) 
        | _ -> x  |> Array.map string |> Array.reduce(fun x y -> x  + " " + y)
    | AGuid(x) ->
        match x.Length with
        | 0 -> "`guid$()"
        | 1 -> "enlist " + (Guid(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + ";" + y) |> (fun x -> "(" + x + ")" )
    | AShort(x) ->
        match x.Length with
        | 0 -> "`short()"
        | 1 -> "enlist " + (Short(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "h" ) 
    | AInt(x) ->
        match x.Length with
        | 0 -> "`int()"
        | 1 -> "enlist " + (Int(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "i")
    | ALong(x) ->
        match x.Length with
        | 0 -> "`long()"
        | 1 -> "enlist " + (Long(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "j")
    | AReal(x) ->
        match x.Length with
        | 0 -> "`real()"
        | 1 -> "enlist " + (Real(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "e")
    | AFloat(x) ->
        match x.Length with
        | 0 -> "`float()"
        | 1 -> "enlist " + (Float(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "f")
    | AChar(x) ->
        match x.Length with
        | 0 -> "`char()"
        | 1 -> "enlist " + (Char(x.[0]) |> stringify )
        | _ -> x |> Array.map string |> Array.reduce(fun x y -> x  +  y)|> (fun x -> "\"" + x + "\"" )
    | AString(x) ->
        match x.Length with
        | 0 -> "`$()"
        | 1 -> "enlist " + (String(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x-> "`"+x) |> Array.reduce(fun x y -> x  + y) 
    | ADate(x) ->
        match x.Length with
        | 0 -> "`date$()"
        | 1 -> "enlist " + (Date(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.dd")) |> Array.reduce(fun x y -> x  + " " + y)
    | ATimestamp(x) ->
        match x.Length with
        | 0 -> "`timestamp$()"
        | 1 -> "enlist " + (Timestamp(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.ddDHH:mm:ss.ffffzzz")) |> Array.reduce(fun x y -> x  + " " + y)
    | AMinute(x) ->
        match x.Length with
        | 0 -> "`minute$()"
        | 1 -> "enlist " + (Minute(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("hh\:mm")) |> Array.reduce(fun x y -> x  + " " + y)
    | ASecond(x) ->
        match x.Length with
        | 0 -> "`second$()"
        | 1 -> "enlist " + (Second(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("hh\:mm\:ss")) |> Array.reduce(fun x y -> x  + " " + y)
    | AMonth(x) ->
        match x.Length with
        | 0 -> "`month$()"
        | 1 -> "enlist " + (Month(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM")) |> Array.reduce(fun x y -> x  + " " + y) |> (fun x -> x + "m")
    | ADateTime(x) ->
        match x.Length with
        | 0 -> "`datetime$()"
        | 1 -> "enlist " + (DateTime(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("yyyy.MM.ddDHH:mm:ss.fff")) |> Array.reduce(fun x y -> x  + " " + y)
    | AKTimespan(x) ->
        match x.Length with
        | 0 -> "`timespan$()"
        | 1 -> "enlist " + (KTimespan(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("hh\:mm\:ss\.fffffff")) |> Array.reduce(fun x y -> x  + " " + y)

    | ATimeSpan(x) ->
        match x.Length with
        | 0 -> "`time$()"
        | 1 -> "enlist " + (TimeSpan(x.[0]) |> stringify )
        | _ -> x |> Array.map (fun x -> x.ToString("hh\:mm\:ss\.fff")) |> Array.reduce(fun x y -> x  + " " + y)
    | Dict(x,y) ->
        let sx = stringify x
        let sy = stringify y
        sx + "!" + sy
    | Flip(x,y) -> 
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


type serialize(n,vt) =
    let mutable j = 0
    let jp() = j<-j+1
    let B = Array.zeroCreate(n)
    let bb b =  if b then 1uy else 0uy
    member x.get() = B
    member x.w(b:bool) =  B.[j]<- b |> bb;jp() 
    member x.w(b:byte)  =  B.[j] <- b;jp()
    member x.w(g:System.Guid)  =
        let a = g.ToByteArray() 
        [| 3; 2; 1; 0; 5; 4; 7; 6; 8; 9; 10; 11; 12; 13; 14; 15 |]
        |> Array.map (fun i  -> a.[i] ) |> Array.iter x.w
    member x.w(b:char)  =  b |> byte |> x.w
    member x.w(b:int16) =  [b;b>>>8] |> List.map byte |> List.iter x.w
    member x.w(b:int)   =  [b;b>>>16] |> List.map int16 |> List.iter x.w
    member x.w(b:int64) =  [b;b>>>32] |> List.map int32 |> List.iter x.w
    member x.w(b:float32) =  b |> System.BitConverter.GetBytes |> Array.iter x.w
    member x.w(b:double) = b |> System.BitConverter.DoubleToInt64Bits |> x.w
    member x.w(b:string) = b |>e.GetBytes |> Array.iter x.w;x.w(0uy)
    member x.w(k:KObject) = 
        k |> t |> byte |> x.w
        let isnull = qn(k)
        match k with
        | Bool(y) -> y |> x.w
        | Guid(y) ->y |> x.w
        | Byte(y) -> y |> x.w
        | Int(y) -> y |> x.w
        | Long(y) -> y |> x.w
        | Short(y) -> y |> x.w
        | Real(y) -> y |> x.w
        | Float(y) -> y |> x.w 
        | Char(y) -> y |> x.w
        | String(y) -> y |> x.w
        | Timestamp(y) -> if vt < 1 then raise(KException("Timestamp not valid pre kdb+2.6"))
                          (if isnull then nj else (y.Ticks - o) * 100L) |> x.w
        | Month(y) -> y |> ftoKMonth|> x.w
        | Date(y) -> y |> ftoKDate |> x.w
        | DateTime(y) -> x.w(if isnull then nf else y.ToOADate() - odate  )
        | KTimespan(y) -> (if isnull then nj else 100L*y.Ticks) |> x.w
        | Minute(y) -> y.Hours*60 + y.Minutes |> x.w
        | Second(y) -> y.Hours*60*60 + y.Minutes*60 + y.Seconds |> x.w
        | TimeSpan(y) -> if vt<1 then raise(KException("Timespan not valid pre kdb+2.6"))
                         x.w(if isnull then ni else (y.Ticks / 10000L |> int))
        | ABool(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AGuid(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AByte(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AShort(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AInt(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | ALong(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AReal(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w
        | AFloat(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AChar(y) -> x.w(0uy);x.w(y.Length);y |> Array.iter x.w 
        | AString(y) -> x.w(0uy);x.w(y.Length);y |> Array.map (fun x -> String(x)) 
                                                 |> Array.iter x.w
        | ATimestamp(y) -> x.w(0uy);x.w(y.Length)
                           y |> Array.map (fun y -> if y.Ticks=nj then nj else ((y.Ticks - o) * 100L) ) 
                             |> Array.iter x.w
        | ADate(y) -> x.w(0uy);x.w(y.Length); y |> Array.map ftoKDate |> Array.iter x.w
        | ADateTime(y) ->x.w(0uy);x.w(y.Length)
                         y |> Array.map (fun y-> y.ToOADate() - odate) 
                           |> Array.iter x.w
        | AKTimespan(y) -> x.w(0uy);x.w(y.Length) 
                           y |> Array.map (fun y -> if y.Ticks = nj then nj else (y.Ticks * 100L)) 
                             |> Array.iter x.w
        | ATimeSpan(y) -> if vt < 1 then raise(KException("Timespan not valid pre kdb+2.6"))
                          x.w(0uy);x.w(y.Length) 
                          y  |> Array.map (fun y -> if ((int)y.Ticks )=ni then ni else (y.Ticks / 10000L |> int)) 
                             |> Array.iter x.w
        | AMonth(y) -> x.w(0uy);x.w(y.Length); y |> Array.map ftoKMonth|> Array.iter x.w
        | AMinute(y) -> x.w(0uy);x.w(y.Length); y|> Array.map (fun y-> y.Hours*60 + y.Minutes )
                                                 |> Array.iter x.w
        | ASecond(y) -> x.w(0uy);x.w(y.Length); y|> Array.map (fun y->y.Hours*60*60+y.Minutes*60+y.Seconds )
                                                 |> Array.iter x.w
        | AKObject(y)-> x.w(0uy);x.w(y.Length); y  |> List.iter x.w
        | Dict(y,z) -> x.w(y); x.w(z);        
        | Flip(y,z) -> x.w(0uy);x.w(99uy);x.w(AString(y)); x.w(AKObject(z));
        | ERROR(x) -> raise(KException(x))
        | NULL -> raise(KException("unknown error"))
        | TODO -> raise(KException("unknown todo"))


    static member wi (i:int,vt) (k:KObject)=
        let n = 8 + nx k
        let B = new serialize(n,vt)
        [1uy;(i |> byte);0uy;0uy] |> List.iter B.w
        B.w(n);B.w(k)
        B.get()


type deserialize(s:System.Net.Sockets.NetworkStream) =
    let maxBufferSize = 65536
    let mutable a = false
    let mutable b = Array.zeroCreate<byte>(1)
    let mutable j = 0
    let jp() = j<-j+1
    let gb() = 
        let r = b.[j]
        jp()
        r
    let rec read_ n i =
        if n > i then
            let a = s.Read(b,i,Math.Min(maxBufferSize,n-i))
            if a=0 then raise(KException("read")) 
            i+a |> read_ n
        else ()

    let rb() = 
        let r = match b.[j] with
                | 0uy -> false
                | 1uy -> true
                | _ -> true // add error
        jp()
        r
    let rx() = gb()

    let rh() = 
        let x = gb() |> int
        let y = gb() |> int
        let r = if a then  (x &&& 0xff) ||| (y <<< 8) else (x <<< 8) ||| (y &&& 0xff)
        r |> int16

    let ri() = 
        let x = rh() |> int
        let y = rh() |> int
        let r = if a then  (x &&& 0xffff) ||| (y <<< 16) else (x <<< 16) ||| (y &&& 0xffff)
        r |> int

    let rj() = 
        let x = ri() |> int64
        let y = ri() |> int64
        let r = if a then  (x &&& 0xffffffffL) ||| (y <<< 32) else (x <<< 32) ||| (y &&& 0xffffffffL)
        r |> int64

    let re() =
        if not(a) then
            let c = b.[j]
            b.[j] <- b.[j+3]
            b.[j+3] <- c
            let c = b.[j+1]
            b.[j+1] <- b.[j+2]
            b.[j+2] <- c

        let r = System.BitConverter.ToSingle(b,j);
        j<-j+4
        r
    let rf() = 
        rj() |> System.BitConverter.Int64BitsToDouble

    let rc() = 
        gb() |> int |> fun x -> x &&& 0xff |> char

    let rs() = 
        let k = j
        let rec r a =
            match a with
            | 0 -> ()
            | _ -> gb() |> int |> r
        gb() |> int |> r
        let s = e.GetString(b,k,j-k-1)
        s

    let rg() =  let oa = a
                a<-false
                let i,h1,h2 =ri(),rh(),rh()
                a <-oa
                let b = Array.init 8 (fun i -> rx())
                new System.Guid(i,h1,h2,b)

    let u()  =
        let mutable n, r, f, s, p, i = 0,0,0,8,8,0s
        j <- 0
        let dst = ri() |> Array.zeroCreate<byte>
        let mutable d = j        
        let aa = Array.zeroCreate<int>(256)
        
        while s < dst.Length do
            if i = 0s then
                f <- 0xff &&& (int)b.[d]
                d<-d+1
                i <- 1s 
            if not ((f &&& (int)i) = 0 ) then
                r <- aa.[0xff &&& (int)b.[d] ]
                d<-d+1
                dst.[s] <- dst.[r]
                s<-s+1;r<-r+1
                dst.[s] <- dst.[r]
                s<-s+1;r<-r+1
                n <- 0xff &&& (int)b.[d];
                d<-d+1
                for m in [0..n-1] do
                     dst.[s + m] <- dst.[r + m]
            else
                 dst.[s] <- b.[d]
                 s<-s+1;d<-d+1
            while p < s-1 do                
                aa.[(0xff &&& (int)dst.[p]) ^^^ (0xff &&& (int)dst.[p + 1])] <- p
                p<-p+1
            
            if (f &&& (int) i) = 0 |> not then
                s<-s+n
                p <-s
            i <- i*2s 
            if i = 256s then i <- 0s
        b <- dst
        j <- 8


    let read n = 
        b <- Array.zeroCreate<byte>(n)
        read_ n 0

    let rec r() = 
        let t = gb() |> int8 |> int

        match t with
        | -1 ->  Bool(rb())
        | -2 ->  Guid(rg())
        | -4 ->  Byte(gb())
        | -5 ->  Short(rh())
        | -6 ->  Int(ri())
        | -7 ->  Long(rj())
        | -8 ->  Real(re())
        | -9 ->  Float(rf())
        | -10 -> Char( rc())
        | -11 -> String( rs())
        | -12 -> rj() |> fun x -> if x<0L then (x+1L)/100L - 1L else x / 100L
                      |> fun x-> Timestamp(new DateTime(x + o))
        | -13 -> let r = ri()
                 match -120000 < r && r < 120000  with
                 | true ->  Month ((new DateTime(2000,1,1)).AddMonths(r))
                 | false -> Month (DateTime.MinValue)
                 //Month ((new DateTime(2000,1,1)).AddMonths(r))
        | -14 -> ri() |> ktofDate |> fun x -> Date(x)
        | -15 -> let r = rf()
                 match System.Double.IsNaN r with
                 | true -> DateTime(DateTime.MinValue)
                 | false -> DateTime(System.DateTime.FromOADate(r + odate))

                 // DateTime(System.DateTime.FromOADate(rf() + odate))
        | -16 -> rj() |> fun x -> KTimespan(new System.TimeSpan(x / 100L))
        | -17 -> let r = ri()
                 Minute(new System.TimeSpan(r /60,r%60,0))
        | -18 -> let r = ri()
                 let s = r % 60
                 let m = ((r-s) / 60) % 60
                 let h = ((((r-s) / 60) - m)/60) % 60
                 Second(new System.TimeSpan(h,m,s))
        | -19 -> ri() |> fun x -> TimeSpan(new System.TimeSpan((x|>int64) * 10000L))
        |  99 -> Dict(r(),r())
        |  98 -> jp(); 
                 match r() with
                 | Dict(x,y) -> match x,y with
                                | AString(x),AKObject(y) -> Flip(x,y)
                                | _,_ -> ERROR("error in kx.r")
                 | _ -> ERROR("error in kx.r")
        |   0 -> jp();AKObject((fun i -> r()) |> List.init (ri())  )
        |   1 -> jp();ABool( (fun i -> rb()) |> Array.init (ri()) )
        |   2 -> jp();AGuid( (fun i -> rg()) |> Array.init (ri()) )
        |   4 -> jp();AByte( (fun i -> gb()) |> Array.init (ri())  )
        |   5 -> jp();AShort( (fun i -> rh()) |> Array.init (ri())  )
        |   6 -> jp();AInt( (fun i -> ri()) |> Array.init (ri())  )
        |   7 -> jp();ALong( (fun i -> rj()) |> Array.init (ri())  )
        | 8 -> jp();AReal( (fun i -> re()) |> Array.init (ri())  )
        | 9 -> jp();AFloat( (fun i -> rf()) |> Array.init (ri())  )
        | 10 -> jp();AChar( (fun i -> rc()) |> Array.init (ri())  ) 
        | 11 -> jp();AString( (fun i -> rs()) |> Array.init (ri()) )
        | 12 -> jp()
                Array.init (ri()) (fun i -> rj())
                |> Array.map (fun x -> if x<0L then (x+1L)/100L - 1L else x / 100L )
                |> Array.map (fun x -> new DateTime(x + o) )
                |> (fun x -> ATimestamp(x))

        | 13 -> jp()
                (fun i -> ri()) 
                |> Array.init (ri()) 
                |> Array.map (fun r -> match -120000 < r && r < 120000  with
                                       | true ->  (new DateTime(2000,1,1)).AddMonths(r)
                                       | false -> DateTime.MinValue)

                |> (fun x -> AMonth(x))

        | 14 -> jp();ADate( (fun i -> ri()) |> Array.init (ri()) |> Array.map ktofDate )
        
        | 15 -> jp()
                (fun i -> rf()) |> Array.init (ri()) 
                                |> Array.map (fun r -> match System.Double.IsNaN r with
                                                       | true -> DateTime.MinValue
                                                       | false -> System.DateTime.FromOADate(r + odate)
                                                       ) 
                                |> (fun x->ADateTime(x))
        | 16 -> jp()
                (fun i -> rj()) |> Array.init (ri())
                                |> Array.map (fun x -> new System.TimeSpan(x / 100L))
                                |> (fun x-> AKTimespan(x))
        | 17 -> jp()
                (fun i -> ri()) 
                |> Array.init (ri()) 
                |> Array.map (fun x ->  new System.TimeSpan(x /60,x % 60,0) )
                |> (fun x -> AMinute(x))
        | 18 -> jp()
                (fun i -> ri()) 
                |> Array.init (ri()) 
                |> Array.map (fun x ->System.TimeSpan(x /(60*60),x%(60*60),x%(60*60*60)) )
                |> (fun x -> ASecond(x))
        | 19 -> jp()
                (fun i -> ri()) |> Array.init (ri())
                                |> Array.map (fun x -> new System.TimeSpan((x|>int64) * 10000L))
                                |> (fun x-> ATimeSpan(x))
        | -128 -> ERROR( rs())
        | 100 -> rs() |> ignore
                 r()
        | t when t>105 -> r()
        | t when t>99 -> if (t = 101 && gb()=0uy) then NULL else ERROR("func")
        | _ -> ERROR("Error in kx.r")

    member x.k() =
        read 8
        a<- 1uy = b.[0]
        let c = 1uy = b.[2]
        j <- 4
        -8 + ri() |> read
        if c then u() else j<-0
        if (int b.[0]) = 128 then ()
        r()

[<AllowNullLiteral>]
type c(h:string,p:int,u:string,maxBufferSize:int) =
    let connect =
        try 
            new System.Net.Sockets.TcpClient(h,p)
        with
            | ex ->
                raise(KException("no_con"))
   
    let s = 
        let s = connect.GetStream()
        // e.GetBytes([|'\003'; '\000'|])|> fun t -> s.Write(t,0,t.Length)
        e.GetBytes(u+"\003"+"\000")|> fun t -> s.Write(t,0,t.Length)
        s
    let vt = 
        let b = Array.zeroCreate<byte>(1)
        let num = s.Read(b,0,1)
        if 1=num |> not then raise(KException("access"))
        3 |> byte |> min b.[0] |> int
    let d = deserialize(s)

    member o.k() = d.k()
    member o.k(x:KObject) = x |> serialize.wi(1,vt) |> fun x -> s.Write(x,0,x.Length)
                            d.k() 
    member o.k(s:string) = AChar(s.ToCharArray()) |> o.k
    member o.k(s:string,k:KObject)  = AKObject([AChar(s.ToCharArray());k]) |> o.k
    member o.k(s:string,k1:KObject,k2:KObject)  = AKObject([AChar(s.ToCharArray());k1;k2]) |> o.k
    member o.k(s:string,k1:KObject,k2:KObject,k3:KObject)  = AKObject([AChar(s.ToCharArray());k1;k2;k3]) |> o.k
    member o.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject)  = AKObject([AChar(s.ToCharArray());k1;k2;k3;k4]) |> o.k
    member o.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject,k5:KObject)  = AKObject([AChar(s.ToCharArray());k1;k2;k3;k4;k5]) |> o.k
    member o.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject,k5:KObject,k6:KObject)  = AKObject([AChar(s.ToCharArray());k1;k2;k3;k4;k5;k6]) |> o.k

    member o.ks(x:KObject) = x |> serialize.wi(0,vt) |> fun x -> s.Write(x,0,x.Length)
    member o.ks(str:string) = AChar(str.ToCharArray()) |> o.ks
    member o.ks(s:string,x:KObject) = AKObject([AChar(s.ToCharArray());x])  |> o.ks
    member o.ks(s:string,x:KObject,y:KObject) = AKObject([AChar(s.ToCharArray());x;y])  |> o.ks
    member o.ks(s:string,x:KObject,y:KObject,z:KObject) = AKObject([AChar(s.ToCharArray());x;y;z])  |> o.ks
    member o.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject) = AKObject([AChar(s.ToCharArray());x;y;z;a])  |> o.ks
    member o.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject,b:KObject) = AKObject([AChar(s.ToCharArray());x;y;z;a;b])  |> o.ks
    member o.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject,b:KObject,c:KObject) = AKObject([AChar(s.ToCharArray());x;y;z;a;b;c])  |> o.ks

    new(h:string,p:int) = new c(h,p,"",65536)
    new(h:string,p:int,u:string) = new c(h,p,u,65536)
    member o.close() = s.Close()
    member o.gvt() = vt
    member o.Connected() = connect.Connected

type cMessages = 
    
    | Init of string*int*string*AsyncReplyChannel<string>
    
    | KMessage0 of string*AsyncReplyChannel<KObject>
    | KMessage1 of string*KObject*AsyncReplyChannel<KObject>
    | KMessage2 of string*KObject*KObject*AsyncReplyChannel<KObject>
    | KMessage3 of string*KObject*KObject*KObject*AsyncReplyChannel<KObject>
    | KMessage4 of string*KObject*KObject*KObject*KObject*AsyncReplyChannel<KObject>
    | KMessage5 of string*KObject*KObject*KObject*KObject*KObject*AsyncReplyChannel<KObject>
    | KMessage6 of string*KObject*KObject*KObject*KObject*KObject*KObject*AsyncReplyChannel<KObject>
    | KMessage7 of AsyncReplyChannel<KObject>
    | KMessage8 of KObject*AsyncReplyChannel<KObject>    

    | KSMessage0 of string
    | KSMessage1 of string*KObject
    | KSMessage2 of string*KObject*KObject
    | KSMessage3 of string*KObject*KObject*KObject
    | KSMessage4 of string*KObject*KObject*KObject*KObject
    | KSMessage5 of string*KObject*KObject*KObject*KObject*KObject
    | KSMessage6 of string*KObject*KObject*KObject*KObject*KObject*KObject
    | KSMessage7 of KObject

type cs() =
    let mutable con:c = null
    let mailbox = MailboxProcessor.Start(fun inbox -> 
        let rec loop () = async {
            let! msg = inbox.Receive()
            match msg with 
            | Init (h,p,u,replyChannel )-> 
                try
                    con <- new c(h,p,u)
                    replyChannel.Reply("connected")
                with
                | KException(x) -> replyChannel.Reply(x)
            | KMessage0(s,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s))

            | KMessage1(s,k1,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1))
                    
            | KMessage2(s,k1,k2,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1,k2))

            | KMessage3(s,k1,k2,k3,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1,k2,k3))

            | KMessage4(s,k1,k2,k3,k4,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1,k2,k3,k4))

            | KMessage5(s,k1,k2,k3,k4,k5,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1,k2,k3,k4,k5))

            | KMessage6(s,k1,k2,k3,k4,k5,k6,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(s,k1,k2,k3,k4,k5,k6))

            | KMessage7(replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k())

            | KMessage8(k1,replyChannel) ->
                match con with
                | null -> replyChannel.Reply(KObject.ERROR("no_con"))
                | _ ->  replyChannel.Reply(con.k(k1))

            | KSMessage0(s) -> 
                match con with
                | null -> ()
                | _ ->  con.ks(s) 

            | KSMessage1(s,k1) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1)
            | KSMessage2(s,k1,k2) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1,k2)

            | KSMessage3(s,k1,k2,k3) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1,k2,k3)

            | KSMessage4(s,k1,k2,k3,k4) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1,k2,k3,k4) 

            | KSMessage5(s,k1,k2,k3,k4,k5) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1,k2,k3,k4,k5)

            | KSMessage6(s,k1,k2,k3,k4,k5,k6) ->
                match con with
                | null -> ()
                | _ ->  con.ks(s,k1,k2,k3,k4,k5,k6)

            | KSMessage7(k1) ->
                match con with
                | null -> ()
                | _ ->  con.ks(k1)                

            return! loop ()
        }
        loop ())

    member this.Init(h:string,p:int,u:string) : Async<string> =
        mailbox.PostAndAsyncReply( fun reply -> Init(h,p,u,reply))
    
    member this.k(s:string) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage0(s,reply))    
    member this.k(s:string,k:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage1(s,k,reply))    
    member this.k(s:string,k1:KObject,k2:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage2(s,k1,k2,reply))    
    member this.k(s:string,k1:KObject,k2:KObject,k3:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage3(s,k1,k2,k3,reply))    
    member this.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage4(s,k1,k2,k3,k4,reply))    
    member this.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject,k5:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage5(s,k1,k2,k3,k4,k5,reply))    
    member this.k(s:string,k1:KObject,k2:KObject,k3:KObject,k4:KObject,k5:KObject,k6:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage6(s,k1,k2,k3,k4,k5,k6,reply))     
     member this.k() : Async<KObject> = 
        mailbox.PostAndAsyncReply( fun reply -> KMessage7(reply))
    member this.k(k:KObject) : Async<KObject> =
        mailbox.PostAndAsyncReply( fun reply -> KMessage8(k,reply))

    member this.ks(x:KObject) =
        mailbox.Post(KSMessage7(x))    
    member this.ks(str:string) =
        mailbox.Post(KSMessage0(str))    
    member this.ks(s:string,x:KObject) =
        mailbox.Post(KSMessage1(s,x))    
    member this.ks(s:string,x:KObject,y:KObject) =
        mailbox.Post(KSMessage2(s,x,y))    
    member this.ks(s:string,x:KObject,y:KObject,z:KObject) =
        mailbox.Post(KSMessage3(s,x,y,z))    
    member this.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject) =
        mailbox.Post(KSMessage4(s,x,y,z,a))    
    member this.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject,b:KObject) =
        mailbox.Post(KSMessage5(s,x,y,z,a,b))    
    member this.ks(s:string,x:KObject,y:KObject,z:KObject,a:KObject,b:KObject,c:KObject,d:KObject) =
        mailbox.Post(KSMessage6(s,x,y,z,a,c,d))    

//let c = kx.c("localhost",8888)
//
//let tf x y = x |> Array.zip y |> Array.fold (fun s (a,b) -> s && a = b) true
//
//let rec m r k =
//    let tl x y =  x |> List.zip y |> List.fold (fun s (a,b)-> s && m a b) true
//    match r,k with
//    | kx.Bool(x),kx.Bool(y) -> x = y
//    | kx.Byte(x),kx.Byte(y) -> x = y
//    | kx.Guid(x),kx.Guid(y) -> x = y    
//    | kx.Short(x),kx.Short(y) -> x = y
//    | kx.Int(x),kx.Int(y) -> x = y
//    | kx.Long(x),kx.Long(y) -> x = y
//    | kx.Real(x),kx.Real(y) -> x = y
//    | kx.Float(x),kx.Float(y) -> x = y
//    | kx.Char(x),kx.Char(y) -> x = y
//    | kx.String(x),kx.String(y) -> x = y
//    | kx.Date(x),kx.Date(y) -> x = y
//    | kx.Month(x),kx.Month(y) -> x = y
//    | kx.Timestamp(x),kx.Timestamp(y) -> x = y
//    | kx.KTimespan(x),kx.KTimespan(y) -> x = y
//    | kx.DateTime(x),kx.DateTime(y) -> x = y
//    | kx.TimeSpan(x),kx.TimeSpan(y) -> x = y
//    | kx.Minute(x),kx.Minute(y) -> x = y
//    | kx.Second(x),kx.Second(y) -> x = y
//    | kx.AKObject(x),kx.AKObject(y) -> x |> tl y
//    | kx.ABool(x),kx.ABool(y) -> x |> tf y
//    | kx.AByte(x),kx.AByte(y) -> x |> tf y
//    | kx.AGuid(x),kx.AGuid(y) -> x |> tf y
//    | kx.AShort(x),kx.AShort(y) -> x |> tf y 
//    | kx.AInt(x),kx.AInt(y) -> x |> tf y
//    | kx.ALong(x),kx.ALong(y) -> x |> tf y
//    | kx.AReal(x),kx.AReal(y) -> x |> tf y
//    | kx.AFloat(x),kx.AFloat(y) -> x |> tf y
//    | kx.AChar(x),kx.AChar(y) -> x |> tf y
//    | kx.AString(x),kx.AString(y) -> x |> tf y
//    | kx.ADate(x),kx.ADate(y) -> x |> tf y
//    | kx.ATimestamp(x),kx.ATimestamp(y) -> x |> tf y
//    | kx.AMinute(x),kx.AMinute(y) -> x |> tf y
//    | kx.ASecond(x),kx.ASecond(y) -> x |> tf y
//    | kx.AMonth(x),kx.AMonth(y) -> x |> tf y
//    | kx.ADateTime(x),kx.ADateTime(y) -> x |> tf y
//    | kx.AKTimespan(x),kx.AKTimespan(y) -> x |> tf y
//    | kx.ATimeSpan(x),kx.ATimeSpan(y) -> x |> tf y    
//    | kx.Dict(x,y),kx.Dict(a,b) -> (x |> m a) && (y |> m b)
//    | kx.Flip(x,y),kx.Flip(a,b) -> (x |> tf a) && (y |> tl b)
//    | x,y -> false
//
//
//let test k =
//    let tf x y = x |> Array.zip y |> Array.fold (fun s (a,b) -> s && a = b) true
//    let r = c.k("upd",k)
//    let comp = m r k
//    match r with
//    | kx.Bool(_) -> printfn "Bool test passed: %b" comp
//    | kx.Byte(_) -> printfn "Byte test passed: %b" comp
//    | kx.Guid(_) -> printfn "Guid test passed: %b" comp    
//    | kx.Short(_) -> printfn "Short test passed: %b" comp
//    | kx.Int(_) -> printfn "Int test passed: %b" comp
//    | kx.Long(_) -> printfn "Long test passed: %b" comp
//    | kx.Real(_) -> printfn "Float test passed: %b" comp
//    | kx.Float(_) -> printfn "Double test passed: %b" comp
//    | kx.Char(_) -> printfn "Char test passed: %b" comp
//    | kx.String(_) -> printfn "String test passed: %b" comp
//    | kx.Date(_) -> printfn "Date test passed: %b" comp
//    | kx.Timestamp(_) -> printfn "Timestamp test passed: %b" comp
//    | kx.Minute(_) -> printfn "Minute test passed: %b" comp
//    | kx.Second(_) -> printfn "Second test passed: %b" comp
//    | kx.Month(_) -> printfn "Month test passed: %b" comp         
//    | kx.DateTime(_) -> printfn "DateTime test passed: %b" comp
//    | kx.TimeSpan(_) -> printfn "TimeSpan test passed: %b" comp                        
//    | kx.KTimespan(_) -> printfn "KTimespan test passed: %b" comp
//    | kx.AKObject(_) -> printfn "AKObject test passed: %b" comp
//    | kx.ABool(_) -> printfn "ABool test passed: %b" comp
//    | kx.AByte(_) -> printfn "AByte test passed: %b" comp
//    | kx.AGuid(_) -> printfn "AGuid test passed: %b" comp
//    | kx.AShort(_) -> printfn "AShort test passed: %b" comp
//    | kx.AInt(_) -> printfn "AInt test passed: %b" comp
//    | kx.ALong(_) -> printfn "ALong test passed: %b" comp
//    | kx.AReal(_) -> printfn "AFloat test passed: %b" comp
//    | kx.AFloat(_) -> printfn "ADouble test passed: %b" comp
//    | kx.AChar(_) -> printfn "AChar test passed: %b" comp
//    | kx.AString(_) -> printfn "AString test passed: %b" comp
//    | kx.ADate(_) -> printfn "ADate test passed: %b" comp
//    | kx.ATimestamp(_) -> printfn "ATimestamp test passed: %b" comp
//    | kx.AMinute(_) -> printfn "AMinute test passed: %b" comp
//    | kx.ASecond(_) -> printfn "ASecond test passed: %b" comp
//    | kx.AMonth(_) -> printfn "AMonth test passed: %b" comp
//    | kx.ADateTime(_) -> printfn "ADateTime test passed: %b" comp
//    | kx.AKTimespan(_) -> printfn "AKTimespan test passed: %b" comp
//    | kx.ATimeSpan(_) -> printfn "TimeSpan test passed: %b" comp                
//    | kx.Dict(_,_) -> printfn "Dict test passed: %b" comp
//    | kx.Flip(_,_) -> printfn "Flip test passed: %b" comp
//    | _ -> printfn "TODO: Add this test for %A" r
//
//[kx.Bool(true);kx.Byte(0uy);kx.Short(1s);
//    kx.Int(5363);kx.Long(1L);kx.Real(3.0 |> float32);kx.Float(3.0);
//    kx.Char('c');kx.String("upd");kx.Date(System.DateTime.ParseExact("2014.09.14","yyyy.MM.dd",null));kx.Timestamp(System.DateTime.ParseExact("2014.09.14 12:12:32.123456","yyyy.MM.dd HH:mm:ss.ffffff",null));kx.Month(new System.DateTime(2014,9,1));kx.Minute(new System.TimeSpan(12,2,0));kx.Second(new System.TimeSpan(12,13,14));kx.DateTime(System.DateTime.ParseExact("2014.09.14 12:12:32.123","yyyy.MM.dd HH:mm:ss.fff",null));kx.KTimespan( new System.TimeSpan(0,12,23,34,345) );kx.Guid(new System.Guid("c8e03a84-8d2a-bf10-cd17-ee98c08910a3") );
//    kx.TimeSpan( new System.TimeSpan(0,18,23,34,345) )
//    ] |> List.iter test
//
//
//[kx.ABool([|true; false|]);kx.AByte([|0uy; 1uy|]);kx.AShort([|1s;2s|]); kx.AInt([|1;2|]);kx.ALong([|1L;2L|]);kx.AReal([|3.0 |> float32;4.0 |> float32|]);kx.AFloat([|3.0;4.0|]);kx.AChar([|'c';'d'|]);kx.AString([|"upd";"upd1"|]);kx.Dict(kx.AString([|"a"|]),kx.AFloat([|3.0|])); kx.Flip([|"ab";"bc"|],[ kx.AFloat([|3.0;4.0|]) ; kx.ABool([|true; false|]) ] );kx.AKObject([ kx.AFloat([|3.0;4.0|]) ; kx.ABool([|true; false|]) ] );
//    kx.ADate([|System.DateTime.ParseExact("2014.09.14","yyyy.MM.dd",null);System.DateTime.ParseExact("2014.09.15","yyyy.MM.dd",null)|]);
//    kx.AMonth([|System.DateTime.ParseExact("2014.09.01","yyyy.MM.dd",null);System.DateTime.ParseExact("2014.10.01","yyyy.MM.dd",null)|]);
//    kx.ADateTime([|System.DateTime.ParseExact("2014.09.14 12:12:32.123","yyyy.MM.dd HH:mm:ss.fff",null);System.DateTime.ParseExact("2014.09.15 12:18:32.123","yyyy.MM.dd HH:mm:ss.fff",null)|]);
//    kx.AKTimespan([|new System.TimeSpan(0,12,23,34,345);new System.TimeSpan(0,12,43,34,845)|]);
//    kx.ATimeSpan([|new System.TimeSpan(0,12,23,34,345);new System.TimeSpan(0,12,43,34,845)|]);
//    kx.AGuid([|new System.Guid("c8e03a84-8d2a-bf10-cd17-ee98c08910a3");new System.Guid("8992d437-a566-f081-cf59-ee6ae1c6ea04")|])
//] |> List.iter test