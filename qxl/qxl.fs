
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
                        con
                  | true -> let con = connectionMaps.Item uid
                            con.close()
                            let con = kx.c(host,port,user)
                            con
        connectionMaps.Add(uid,con)
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
                    
    // Helper that will create a timer that ticks at timerInterval for timerDuration, then stops
    // Not exported to Excel (incompatible type)
    let createTimer timerInterval timerDuration =
        // setup a timer
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
        timer.Elapsed |> Observable.map (fun elapsed -> [|"H: "+DateTime.Now.ToString("HH:mm:ss.fff");"G: "+DateTime.Now.ToString("HH:mm:ss.fff")|] |> Array.map(fun x -> x :> obj) :> obj) 

    // Excel function to start the timer - using the fact that F# events implement IObservable
    [<ExcelFunction(Description="start timer")>]
    let startTimer timerInterval timerDuration =
        FsAsyncUtil.excelObserve "startTimer" [|float timerInterval; float timerDuration|] (createTimer timerInterval timerDuration)


    type kdbsubscriber(host:string,port:int,user:string) =
        let host ,port, user=host ,port, user
        let con = kx.c(host,port,user)
        let on_event = new Event<kx.KObject>()

        member o.onEvent = on_event.Publish
        member o.start() =
            async {
                while true do
                    let r = con.k()
                    on_event.Trigger(r)
            } |> Async.Start 
        member o.sub(table:string,sym:string) = 
            con.ks(".u.sub",kx.KObject.String(table),kx.KObject.String(sym)) 
            o.start()
        member o.k() = con.k()
            

    let subscriberMaps = new Dictionary<string,kdbsubscriber>()

    [<ExcelFunction(Description="Create subscriber")>]
    let open_subscriber (uid:string) (host:string) (port:int) (user:string) =
        let uid = match subscriberMaps.ContainsKey uid with
                  | false -> 
                        let con = kdbsubscriber(host,port,user)
                        subscriberMaps.Add(uid,con)
                        uid
                  | true -> uid
    
        uid

    [<ExcelFunction(Description="subscribe to table and sym")>]
    let subscribe (uid:string) (table:string) (sym:string) =
        let s = subscriberMaps.Item uid
        s.sub(table,sym)
        FsAsyncUtil.excelObserve "subscribe" [|uid; table;sym|] (s.onEvent |> Observable.map (fun x -> 1 :> obj) )

    [<ExcelFunction(Description="call manually")>]
    let mcall (uid:string) =
        let s = subscriberMaps.Item uid
        s.k() |> ktox 
    
        


