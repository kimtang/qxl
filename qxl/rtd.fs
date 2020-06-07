namespace rtd

module RtdClock = 
    open System
    open System.Collections.Generic
    open System.Runtime.InteropServices
    open System.Threading
    open ExcelDna.Integration.Rtd
    
    // [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    // [ProgId(RtdClockServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)    
    [<ProgId("rtdclock.rtdclockserver");ComVisible(true)>]
    type RtdClockServer(timer,topics) =
        inherit ExcelRtdServer()

        let mutable _timer : Timer = timer
        let mutable _topics : List<ExcelRtdServer.Topic> = topics

        static member ServerProgId = "rtdclock.rtdclockserver"

        member this.timer_tick(_unused_state_:Object) = 
            let now = DateTime.Now.ToString("HH:mm:ss") :> obj          
            _topics.ForEach(fun x -> x.UpdateValue(now) )

        new() = RtdClockServer(null,null)
           
        override this.ServerStart() =
            _timer <- new Timer(this.timer_tick,null,0,1000)
            _topics <- new List<ExcelRtdServer.Topic>()
            true
        override this.ServerTerminate() = 
            _timer.Dispose()

        override this.ConnectData(topic:ExcelRtdServer.Topic, topicInfo:IList<string>, newValues:byref<bool>) = 
            // this._topics <- topic |> List.append this._topics 
            _topics.Add topic
            DateTime.Now.ToString("HH:mm:ss") + " (ConnectData)" :> obj
        
        override this.DisconnectData(topic:ExcelRtdServer.Topic) = 
            _topics.Remove topic |> ignore


module RtdKdb = 
    open System
    open System.Collections.Generic
    open System.Runtime.InteropServices
    open System.Threading
    open ExcelDna.Integration.Rtd

        // dict Seq.empty<string*kx.c>

    type kdb_subscriber(uid:string,host:string,port:int,user:string,sub:string) = 
        let uid,host,port,user,sub = uid,host,port,user,sub
        let mutable connection:Option<kx.c> = None // kx.c(host,port,user)
        let mutable callback:(string*kx.KObject -> unit) = (fun (x,y) -> ())
        let mutable inSub = false
        
        let agent = MailboxProcessor<AsyncReplyChannel<string*kx.KObject>>.Start(fun inbox ->
            let rec loop () =
                async {
                        let! replyChannel = inbox.Receive();
                        // Delay so that the responses come in a different order.
                        //do! Async.Sleep( 5000 - 1000 * n);
                        let o = connection.Value.k()
                        let o = match o with
                                | kx.KObject.AKObject(x) -> x.[2]
                                | _ -> o
                        replyChannel.Reply((uid,o))
                        do! loop ()
                }
            loop ()
            )

        let rec outerloop() =
            let messageAsync = agent.PostAndAsyncReply(fun replyChannel -> replyChannel)
        
            Async.StartWithContinuations(messageAsync, 
                 (fun (id,reply) -> callback((id,reply))
                                    outerloop()
                 ),
                 (fun _ -> ()),
                 (fun _ -> ()))

        member this.Callback
            with set(c) = callback <- c

        member this.InSub
            with get() = inSub

        member this.Uid = uid
        member this.loop() = if inSub then ()
                             else
                                connection <- Some(kx.c(host,port,user))
                                connection.Value.ks(sub)
                                outerloop()
        member this.Desc 
            with get() = (uid,host,port,user,sub)
        member this.Connected =
            match connection with
            | Some(c) -> c.Connected()
            | None -> false

        member this.Close() =
            match connection with
            | Some(c) -> c.close()
            | None -> ()


    // let subscriberMaps = new Dictionary<string,kx.c>()
    let subscriberMaps = Dictionary<string,kdb_subscriber>()
    let objMaps = Dictionary<string,obj>()

    [<ProgId("rtdkdb.rtdkdbserver");ComVisible(true)>]
    type RtdKdbServer(topics) =
        inherit ExcelRtdServer()

        let mutable _topics : (string*ExcelRtdServer.Topic) list = topics
        let mutable seq = 0

        new() = RtdKdbServer(List.empty<string*ExcelRtdServer.Topic>)

        member this.receiveData(id:string,o:kx.KObject) =
            seq <- seq + 1
            match objMaps.ContainsKey id with
            | true -> objMaps.[id] <- ktox.ktox o
            | false -> objMaps.Add(id,o |> ktox.ktox)
            
            _topics |> List.filter(fun (id0,_) -> id0 = id)|>List.iter(fun (_,topic) -> topic.UpdateValue(seq :> obj) )
       
        override this.ServerStart() =
            true // nothing to do 

        override this.ServerTerminate() = 
            // nothing to do
            ()

        override this.ConnectData(topic:ExcelRtdServer.Topic, topicInfo:IList<string>, newValues:byref<bool>) =
            let id = topicInfo.[0]
            let r = match subscriberMaps.ContainsKey id with
                    | true -> _topics <- [(id,topic)] |> List.append _topics
                              let c = subscriberMaps.[id]
                              c.Callback<- this.receiveData
                              c.loop()
                              DateTime.Now.ToString("HH:mm:ss") + " (ConnectData)" :> obj
                    | false -> DateTime.Now.ToString("HH:mm:ss") + " (topic not found)" :> obj
            r
            
    
        override this.DisconnectData(topic:ExcelRtdServer.Topic) =
            let ids = _topics |> List.filter ( fun (id,topic0) -> topic0 = topic )
                              |> List.map (fun (id,_) -> id) |> Seq.distinct
            _topics <- _topics |> List.filter ( fun (id,topic0) -> topic0 = topic |> not )
            
            ids |> Seq.iter(fun id -> 
                                match subscriberMaps.ContainsKey id with
                                | true -> let c = subscriberMaps.[id]
                                          let uid,host,port,user,sub = c.Desc
                                          let b = kdb_subscriber(uid,host,port,user,sub)
                                          subscriberMaps.[id] <- b
                                | false -> ()
                           )
            
            ()
            // _topics.Remove topic |> ignore