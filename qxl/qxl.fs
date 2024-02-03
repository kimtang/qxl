
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

    // we want to turn a matrix either in dic or tbl
    // let's check
    let xtok2dim (arg: obj[,]) =
        let r = arg[0,*] |> xtok1dim
        let c = arg[*,0] |> xtok1dim
        let rcnt = arg |> Array2D.length1
        let lcnt = arg |> Array2D.length2
        
        match fst r,fst c with
        | XString,_ -> 
                        let a = [0 .. lcnt - 1] |> List.map (fun i -> xtok1dim arg[1 .. ,i]) |> List.map snd 
                        let s = match (r  |> snd) with
                                | kx.KObject.AString(s0) -> s0
                                | _ -> Array.create rcnt "ExcelEmpty"
                        kx.Flip(s, a)
        | _,XString -> 
                        let a = [0 .. rcnt - 1] |> List.map (fun i -> xtok1dim arg[i,1 .. ]) |> List.map snd |> kx.KObject.AKObject
                        kx.Dict(snd c, a)
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


    [<ExcelFunction(Description="async execute query")>]
    let dna_aopen_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)= 
        FsAsyncUtil.excelRunAsync "async_open_connection" [|uid :> obj; host :> obj; port :> obj; user :> obj ; passwd :> obj |] (async_open_connection uid host port user passwd )

    [<ExcelFunction(Description="execute query")>]
    let dna_execute (random:obj) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 
        match connectionMaps.ContainsKey uid with
        | false -> "uid not found" :> obj
        | true -> let con = connectionMaps.Item uid
                  let x,y,z,a,b = xtok x,xtok y,xtok z,xtok a,xtok b
                  
                  let r = match x,y,z,a,b with
                          | ExcelMissing,_,_,_,_-> con.k(query)
                          | KObject(x),ExcelMissing,_,_,_-> con.k(query,x)
                          | KObject(x),KObject(y),ExcelMissing,_,_-> con.k(query,x,y)
                          | KObject(x),KObject(y),KObject(z),ExcelMissing,_-> con.k(query,x,y,z)
                          | KObject(x),KObject(y),KObject(z),KObject(a),ExcelMissing-> con.k(query,x,y,z,a)
                          | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b)-> con.k(query,x,y,z,a,b)
                          // | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),KObject(c)-> con.k(query,x,y,z,a,b,c)
                  ktox r

    let asyncExecute (random) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) =
        async {
                match csMaps.ContainsKey uid with
                | false -> return "uid not found" :> obj
                | true -> let con = csMaps.Item uid
                          let x,y,z,a,b = xtok x,xtok y,xtok z,xtok a,xtok b                  
                          let! r = match x,y,z,a,b with
                                  | ExcelMissing,_,_,_,_-> con.k(query)
                                  | KObject(x),ExcelMissing,_,_,_-> con.k(query,x)
                                  | KObject(x),KObject(y),ExcelMissing,_,_-> con.k(query,x,y)
                                  | KObject(x),KObject(y),KObject(z),ExcelMissing,_-> con.k(query,x,y,z)
                                  | KObject(x),KObject(y),KObject(z),KObject(a),ExcelMissing-> con.k(query,x,y,z,a)
                                  | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b)-> con.k(query,x,y,z,a,b)
                                  // | KObject(x),KObject(y),KObject(z),KObject(a),KObject(b),KObject(c)-> con.k(query,x,y,z,a,b,c)
                          return ktox r
        }

    [<ExcelFunction(Description="async execute query")>]
    let dna_aexecute (random:obj) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 
        FsAsyncUtil.excelRunAsync "asyncExecute" [|random; (uid :> obj); (query :> obj); x; y; z; a; b |] (asyncExecute random uid query x y z a b )
        

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

    [<ExcelFunction(Description="Provides a ticking clock")>]
    let dna_rtd(progId:string) (server:string) (topic:string)=
        XlCall.RTD(progId,server,topic)