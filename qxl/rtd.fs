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
            _timer <- new Timer(this.timer_tick,null,0,7000)
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

    type kdb_subscriber(uid:string,host:string,port:int,user:string,sub:string,append:int) = 
        let uid,host,port,user,sub,append = uid,host,port,user,sub,append
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
            with get() = (uid,host,port,user,sub,append)
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
    let kobjArray = Dictionary<string,kx.KObject array>()
    let kobjMaps = Dictionary<string,kx.KObject>()

    [<ProgId("rtdkdb.rtdkdbserver");ComVisible(true)>]
    type RtdKdbServer(topics) =
        inherit ExcelRtdServer()

        let mutable _topics : (string*ExcelRtdServer.Topic) list = topics
        let mutable seq = 0

        new() = RtdKdbServer(List.empty<string*ExcelRtdServer.Topic>)

        member this.receiveData(id:string,o:kx.KObject) =

            let _,_,_,_,_,append = match subscriberMaps.ContainsKey id with
                                   | true -> subscriberMaps.[id].Desc
                                   | false -> "  ","  ",1,"  ","  ",1


            let mergeFlip1 (nkobject:kx.KObject,okobject:kx.KObject) = 
                match nkobject,okobject with
                | kx.AKObject(x),kx.AKObject(y) -> kx.AKObject(x |> List.append y |> (fun x -> x|> List.take (min x.Length append) )   )
                | kx.ABool(x),kx.ABool(y) -> kx.ABool(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ) )
                | kx.AByte(x),kx.AByte(y) -> kx.AByte(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ) )
                | kx.AGuid(x),kx.AGuid(y) -> kx.AGuid(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ) )
                | kx.AShort(x),kx.AShort(y) -> kx.AShort(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) )) 
                | kx.AInt(x),kx.AInt(y) -> kx.AInt(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.ALong(x),kx.ALong(y) -> kx.ALong(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AReal(x),kx.AReal(y) -> kx.AReal(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AFloat(x),kx.AFloat(y) -> kx.AFloat(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AChar(x),kx.AChar(y) -> kx.AChar(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AString(x),kx.AString(y) -> kx.AString(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ) )
                | kx.ADate(x),kx.ADate(y) -> kx.ADate(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.ATimestamp(x),kx.ATimestamp(y) -> kx.ATimestamp(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AMinute(x),kx.AMinute(y) -> kx.AMinute(x |> Array.append y|> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.ASecond(x),kx.ASecond(y) -> kx.ASecond(x |> Array.append y|> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AMonth(x),kx.AMonth(y) -> kx.AMonth(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.ADateTime(x),kx.ADateTime(y) -> kx.ADateTime(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.AKTimespan(x),kx.AKTimespan(y) -> kx.AKTimespan(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | kx.ATimeSpan(x),kx.ATimeSpan(y) -> kx.ATimeSpan(x |> Array.append y |> (fun x -> x|> Array.take (min x.Length append) ))
                | _ -> nkobject

            let mergeFlip (nkobject:kx.KObject) (okobject:kx.KObject) =
                match nkobject,okobject with
                | kx.Flip(ox,oy),kx.Flip(nx,ny) -> let sameHeader = nx |> Array.zip ox |> Array.map (fun (x,y) -> x=y ) |> Array.reduce (fun x y -> x && y) 
                                                   match sameHeader with
                                                   | false -> None
                                                   | true -> let nk = oy |> List.zip ny
                                                                         |> List.map mergeFlip1
                                                             Some(kx.Flip(nx,nk))
                | _ -> None
                
            seq <- seq + 1

            let nkobject = match append,kobjMaps.ContainsKey id with
                           | _,false ->  kobjMaps.Add(id,o)
                                         o
                           | 0,true ->   kobjMaps.[id] <- o                           
                                         o
                           | _,true -> let oldkobject =  kobjMaps.[id]                                     
                                       let mo = mergeFlip o oldkobject
                                       match mo with
                                       | None -> kobjMaps.[id]<-o
                                                 o
                                       | Some(mo) -> kobjMaps.[id]<-mo
                                                     mo

            let o1 = ktox.ktox nkobject
            match objMaps.ContainsKey id with
            | true -> objMaps.[id] <- o1
            | false -> objMaps.Add(id,o1)
            
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
                                          let uid,host,port,user,sub,append = c.Desc
                                          let b = kdb_subscriber(uid,host,port,user,sub,append)
                                          subscriberMaps.[id] <- b
                                | false -> ()
                           )
            
            ()
            // _topics.Remove topic |> ignore