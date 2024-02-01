
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

    type XType = | XExcelEmpty | XExcelError | XExcelMissing | XBool | XString | XFloat

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
                | _,1  ->   let r= o[*,0] |> Array.map ( fun y -> 
                                                            match y with
                                                            | :? ExcelEmpty -> XExcelEmpty 
                                                            | :? ExcelError -> XExcelError 
                                                            | :? ExcelMissing -> XExcelMissing
                                                            | :? bool as o-> XBool
                                                            | :? string as o-> XString
                                                            | :? float as o-> XFloat
                                                            ) |> Array.distinct |> Array.length
                            match r with
                            | 1 ->  let oo = o[0,0]
                                    let r = match oo with
                                            | :? ExcelEmpty -> Array.create rcnt "ExcelEmpty" |> kx.KObject.AString
                                            | :? ExcelError -> Array.create rcnt "ExcelError" |> kx.KObject.AString
                                            | :? ExcelMissing -> Array.create rcnt "ExcelMissing" |> kx.KObject.AString 
                                            | :? bool -> o[*,0] |> Array.map (fun x -> x :?> bool) |> kx.KObject.ABool
                                            | :? string -> o[*,0] |> Array.map (fun x -> x :?> string) |> kx.KObject.AString
                                            | :? float -> o[*,0] |> Array.map (fun x -> x :?> float) |> kx.KObject.AFloat
                                    KObject(r)
                            | _ -> let r = o[*,0] |> Array.map ( fun y -> 
                                                            match y with
                                                            | :? ExcelEmpty -> kx.KObject.String("ExcelEmpty")
                                                            | :? ExcelError -> kx.KObject.String("ExcelError")
                                                            | :? ExcelMissing -> kx.KObject.String("ExcelMissing")
                                                            | :? bool as o-> kx.KObject.Bool(o)
                                                            | :? string as o-> kx.KObject.String(o)
                                                            | :? float as o-> kx.KObject.Float(o)
                                                            )
                                                  |> Array.toList
                                                  |> kx.KObject.AKObject
                                   KObject(r)
                | 1,_  ->   let rr= o[0,*] 
                            let r = rr |> Array.map ( fun y -> 
                                                            match y with
                                                            | :? ExcelEmpty -> XExcelEmpty 
                                                            | :? ExcelError -> XExcelError 
                                                            | :? ExcelMissing -> XExcelMissing
                                                            | :? bool as o-> XBool
                                                            | :? string as o-> XString
                                                            | :? float as o-> XFloat
                                                            ) |> Array.distinct |> Array.length
                            match r with
                            | 1 ->  let oo = o[0,0]
                                    let r = match oo with
                                            | :? ExcelEmpty -> Array.create rcnt "ExcelEmpty" |> kx.KObject.AString
                                            | :? ExcelError -> Array.create rcnt "ExcelError" |> kx.KObject.AString
                                            | :? ExcelMissing -> Array.create rcnt "ExcelMissing" |> kx.KObject.AString 
                                            | :? bool -> o[0,*] |> Array.map (fun x -> x :?> bool) |> kx.KObject.ABool
                                            | :? string -> o[0,*] |> Array.map (fun x -> x :?> string) |> kx.KObject.AString
                                            | :? float -> o[0,*] |> Array.map (fun x -> x :?> float) |> kx.KObject.AFloat
                                    KObject(r)
                            | _ -> let r = o[0,*] |> Array.map ( fun y -> 
                                                            match y with
                                                            | :? ExcelEmpty -> kx.KObject.String("ExcelEmpty")
                                                            | :? ExcelError -> kx.KObject.String("ExcelError")
                                                            | :? ExcelMissing -> kx.KObject.String("ExcelMissing")
                                                            | :? bool as o-> kx.KObject.Bool(o)
                                                            | :? string as o-> kx.KObject.String(o)
                                                            | :? float as o-> kx.KObject.Float(o)
                                                            )
                                                  |> Array.toList
                                                  |> kx.KObject.AKObject
                                   KObject(r)
                | _,_  ->
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


    let async_open_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)=
        async {
            let uid = match port with
                      | 0 -> ExcelMissing :> obj
                      | _ ->
                             match connectionMaps.ContainsKey uid with
                             | false -> 
                                          let con = kx.c(host,port,user+":"+passwd)
                                          connectionMaps.Add(uid,con)
                                          uid :> obj
                             | true -> let con = connectionMaps.Item uid
                                       match con.Connected() with
                                       | true -> uid :> obj
                                                 
                                       | false -> connectionMaps.Remove uid |> ignore
                                                  let con = kx.c(host,port,user+":"+passwd)
                                                  connectionMaps.Add(uid,con)
                                                  uid :> obj      
            return uid
        }


    [<ExcelFunction(Description="async execute query")>]
    let dna_aopen_connection (uid:string) (host:string) (port:int) (user:string) (passwd:string)= 
        FsAsyncUtil.excelObserve "async_open_connection" [|uid :> obj; host :> obj; port :> obj; user :> obj ; passwd :> obj |]
            (FsAsyncUtil.observeAsync (async_open_connection uid host port user passwd ))

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
            let r = 
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
            return r
        }

    [<ExcelFunction(Description="async execute query")>]
    let dna_axecute (random:obj) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 
        FsAsyncUtil.excelObserve "asyncExecute" [|random; (uid :> obj); (query :> obj); x; y; z; a; b |]
            (FsAsyncUtil.observeAsync(asyncExecute random uid query x y z a b ))
        

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