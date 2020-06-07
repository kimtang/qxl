
namespace qxl

open System
open System.Threading
open ExcelDna.Integration

/// Some utility functions for connecting Excel-DNA async with F#
module FsAsyncUtil =
    /// A helper to pass an F# Async computation to Excel-DNA 
    let excelRunAsync functionName parameters async =
        let obsSource =
            ExcelObservableSource(
                fun () -> 
                { new IExcelObservable with
                    member __.Subscribe observer =
                        // make something like CancellationDisposable
                        let cts = new CancellationTokenSource ()
                        let disp = { new IDisposable with member __.Dispose () = cts.Cancel () }
                        // Start the async computation on this thread
                        Async.StartWithContinuations 
                            (   async, 
                                ( fun result -> 
                                    observer.OnNext(result)
                                    observer.OnCompleted () ),
                                ( fun ex -> observer.OnError ex ),
                                ( fun ex ->
                                    observer.OnCompleted () ),
                                cts.Token 
                            )
                        // return the disposable
                        disp
                }) 
        ExcelAsyncUtil.Observe (functionName, parameters, obsSource)

    /// A helper to pass an F# IObservable to Excel-DNA
    let excelObserve functionName parameters observable = 
        let obsSource =
            ExcelObservableSource(
                fun () -> 
                { new IExcelObservable with
                    member __.Subscribe observer =
                        // Subscribe to the F# observable
                        Observable.subscribe (fun value -> observer.OnNext (value)) observable
                })
        ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
            

module qxl =
    open System.Collections.Generic
    open kx
    open ktox

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
    let dna_open_connection (uid:string) (host:string) (port:int) (user:string) =
        match port with
        | 0 -> ExcelMissing :> obj
        | _ ->

            let con =   match connectionMaps.ContainsKey uid with
                        | false ->  let con = kx.c(host,port,user)
                                // | ex -> raise (new XlCallException(XlCall.XlReturn.XlReturnFailed)
                                    connectionMaps.Add(uid,con)
                                    con

                        | true ->   let con = connectionMaps.Item uid
                                    match con.Connected() with
                                    | true -> con
                                    | false ->  connectionMaps.Remove uid |> ignore
                                                let con = kx.c(host,port,user)
                                                connectionMaps.Add(uid,con)
                                                con
        
            uid :> obj


    let async_open_connection (uid:string) (host:string) (port:int) (user:string) =
        let on_event = new Event<obj>()
        async {
            let uid = match port with
                      | 0 -> ExcelMissing :> obj
                      | _ ->
                             match connectionMaps.ContainsKey uid with
                             | false -> 
                                          let con = kx.c(host,port,user)
                                          connectionMaps.Add(uid,con)
                                          uid :> obj
                             | true -> let con = connectionMaps.Item uid
                                       match con.Connected() with
                                       | true -> uid :> obj
                                                 
                                       | false -> connectionMaps.Remove uid |> ignore
                                                  let con = kx.c(host,port,user)
                                                  connectionMaps.Add(uid,con)
                                                  uid :> obj      

            on_event.Trigger(uid)
        } |> Async.Start
        on_event.Publish

    [<ExcelFunction(Description="async execute query")>]
    let dna_aopen_connection (uid:string) (host:string) (port:int) (user:string) = 
        FsAsyncUtil.excelObserve "async_open_connection" [|uid :> obj; host :> obj; port :> obj; user :> obj |] (async_open_connection uid host port user )

    [<ExcelFunction(Description="execute query")>]
    let dna_execute (random:float) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 
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
        let on_event = new Event<obj>()
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
            on_event.Trigger(r)
        } |> Async.Start
        on_event.Publish 

    [<ExcelFunction(Description="async execute query")>]
    let dna_axecute (random:float) (uid:string) (query:string) (x:obj) (y:obj) (z:obj) (a:obj) (b:obj) = 
        FsAsyncUtil.excelObserve "asyncExecute" [|(random :> obj); (uid :> obj); (query :> obj); x; y; z; a; b |] (asyncExecute random uid query x y z a b )
        

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
            do! Async.Sleep timerDuration
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
    let dna_open_subscriber (uid:string) (host:string) (port:int) (user:string) (sub:string)=
        match port with
        | 0 -> ExcelMissing :> obj
        | _ ->
            match rtd.RtdKdb.subscriberMaps.ContainsKey uid with
            | false ->
                  let con = rtd.RtdKdb.kdb_subscriber(uid,host,port,user,sub)
                  rtd.RtdKdb.subscriberMaps.Add(uid,con)               
            | true ->
                  rtd.RtdKdb.subscriberMaps.[uid].Close()
                  rtd.RtdKdb.subscriberMaps.[uid] <- rtd.RtdKdb.kdb_subscriber(uid,host,port,user,sub)
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
        
    
        


