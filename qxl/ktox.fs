
module ktox

open kx

let conv_as_obj0(k:KObject) = 
    match k with
    | ERROR -> raise (KException("Error in conv_as_obj0"))
    | TODO -> raise (KException("TODO in conv_as_obj0"))
    | NULL -> raise (KException("Null in conv_as_obj0"))
    | Bool(x) -> x :> obj
    | Guid(x) -> x |> string :> obj
    | Byte(x) -> x |> float :> obj
    | Short(x) -> x |> float :> obj 
    | Int(x) -> x |> float :> obj
    | Long(x) -> x |> float :> obj
    | Real(x) -> x |> float :> obj
    | Float(x) -> x :> obj
    | Char(x) -> x |> string :> obj
    | String(x) -> "`"+x :> obj 
    | Timestamp(x) -> x :> obj
    | Month(x) -> x :> obj
    | Date(x) -> x :> obj
    | DateTime(x) -> x :> obj
    | KTimespan(x) -> x.TotalDays :> obj
    | Minute(x) -> x.TotalDays :> obj
    | Second(x) -> x.TotalDays :> obj
    | TimeSpan(x) -> x.TotalDays :> obj
    | ABool(x) ->  ABool(x)  |> stringify :> obj
    | AByte(x) -> AByte(x) |> stringify :> obj
    | AGuid(x) -> AGuid(x) |> stringify :> obj
    | AShort(x) -> AShort(x) |> stringify :> obj
    | AInt(x) -> AInt(x) |> stringify :> obj
    | ALong(x) -> ALong(x) |> stringify :> obj
    | AReal(x) -> AReal(x) |> stringify :> obj
    | AFloat(x) -> AFloat(x) |> stringify :> obj
    | AChar(x) -> AChar(x) |> stringify :> obj
    | AString(x) -> AString(x) |> stringify :> obj
    | ADate(x) -> ADate(x) |> stringify :> obj
    | ATimestamp(x) -> ATimestamp(x) |> stringify :> obj
    | AMinute(x) -> AMinute(x) |> stringify :> obj
    | ASecond(x) -> ASecond(x) |> stringify :> obj
    | AMonth(x) -> AMonth(x) |> stringify :> obj
    | ADateTime(x) -> ADateTime(x) |> stringify :> obj
    | AKTimespan(x) -> AKTimespan(x) |> stringify :> obj
    | ATimeSpan(x) -> ATimeSpan(x) |> stringify :> obj
    | Dict(x,y) -> Dict(x,y) |> stringify :> obj
    | Flip(x,y) -> Flip(x,y) |> stringify :> obj
    | AKObject(x) ->  AKObject(x)  |> stringify :> obj


let conv_as_array1(k:KObject) = 
    match k with
    | Bool(x) -> x :> obj |> Array.create 1
    | Guid(x) -> x |> string :> obj |> Array.create 1
    | Byte(x) -> x |> float :> obj |> Array.create 1
    | Short(x) -> x |> float :> obj |> Array.create 1
    | Int(x) -> x |> float :> obj |> Array.create 1
    | Long(x) -> x |> float :> obj |> Array.create 1
    | Real(x) -> x |> float :> obj |> Array.create 1
    | Float(x) -> x :> obj |> Array.create 1
    | Char(x) -> x |> string :> obj |> Array.create 1
    | String(x) -> "`"+x :> obj |> Array.create 1
    | Timestamp(x) -> x :> obj |> Array.create 1
    | Month(x) -> x :> obj |> Array.create 1
    | Date(x) -> x :> obj |> Array.create 1
    | DateTime(x) -> x :> obj |> Array.create 1
    | KTimespan(x) -> x.TotalDays :> obj |> Array.create 1
    | Minute(x) -> x.TotalDays :> obj |> Array.create 1
    | Second(x) -> x.TotalDays :> obj |> Array.create 1
    | TimeSpan(x) -> x.TotalDays :> obj |> Array.create 1
        
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
    | AKTimespan(x) -> x |> Array.map (fun x -> x.TotalDays :> obj) 
    | AMinute(x) -> x |> Array.map (fun x -> x.TotalDays :> obj) 
    | ASecond(x) -> x |> Array.map (fun x -> x.TotalDays :> obj) 
    | ATimeSpan(x) -> x |> Array.map (fun x -> x.TotalDays :> obj) 
    | AKObject(x) -> x |> List.map conv_as_obj0 |> List.toArray
    | Dict(x,y) -> Dict(x,y) |> stringify :> obj |> Array.create 1 
    | Flip(x,y) -> Flip(x,y) |> stringify :> obj |> Array.create 1 
    | _ -> raise (KException("Error in conv_as_array1"))


let rec conv_as_matrix0(k:KObject) = 
    match k with    
    | ABool(x) -> ABool(x) |> conv_as_array1 |> Array.create 1
    | AGuid(x) -> AGuid(x) |> conv_as_array1 |> Array.create 1
    | AByte(x) -> AByte(x) |> conv_as_array1 |> Array.create 1
    | AShort(x) -> AShort(x) |> conv_as_array1 |> Array.create 1
    | AInt(x) -> AInt(x) |> conv_as_array1 |> Array.create 1
    | ALong(x) -> ALong(x) |> conv_as_array1 |> Array.create 1
    | AReal(x) -> AReal(x) |> conv_as_array1 |> Array.create 1
    | AFloat(x) -> AFloat(x) |> conv_as_array1 |> Array.create 1
    | AChar(x) -> AChar(x) |> conv_as_array1 |> Array.create 1
    | AString(x) -> AString(x) |> conv_as_array1 |> Array.create 1
    | ATimestamp(x) -> ATimestamp(x) |> conv_as_array1 |> Array.create 1
    | AMonth(x) -> AMonth(x) |> conv_as_array1 |> Array.create 1
    | ADate(x) -> ADate(x) |> conv_as_array1 |> Array.create 1
    | ADateTime(x) -> ADateTime(x) |> conv_as_array1 |> Array.create 1
    | AKTimespan(x) -> AKTimespan(x) |> conv_as_array1 |> Array.create 1
    | AMinute(x) -> AMinute(x) |> conv_as_array1 |> Array.create 1
    | ASecond(x) -> ASecond(x) |> conv_as_array1 |> Array.create 1
    | ATimeSpan(x) -> ATimeSpan(x) |> conv_as_array1 |> Array.create 1
    | AKObject(x) -> AKObject(x) |> conv_as_array1 |> Array.create 1

    | Dict(x,y) -> // hmmm, again a dict in 
        let xc = conv_as_matrix0 x
        let yc = conv_as_matrix0 y
        let yc = match yc.Length with
                 | 1 -> yc
                 | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
        let r = yc |> Array.append xc 
        r
    | Flip(x,y) ->
        let xs = x |> Array.map (fun x -> [|x :> obj|]) 
        let yc = y |> List.map conv_as_array1  |> List.toArray
        // get the size an merge it together
        let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
        r
    | _ -> raise (KException("Error in conv_as_matrix0"))


let ktox( k:KObject) =
    match k with
    | ERROR -> "Error in ktox" :> obj
    | TODO -> raise (KException("TODO in ktox"))
    | NULL -> raise (KException("Null in ktox"))
    | Bool(x) -> Bool(x)|> conv_as_obj0
    | Guid(x) -> Guid(x)|> conv_as_obj0
    | Byte(x) -> Byte(x)|> conv_as_obj0
    | Short(x) -> Short(x)|> conv_as_obj0
    | Int(x) -> Int(x)|> conv_as_obj0
    | Long(x) -> Long(x)|> conv_as_obj0
    | Real(x) -> Real(x)|> conv_as_obj0
    | Float(x) -> Float(x)|> conv_as_obj0
    | Char(x) -> Char(x)|> conv_as_obj0
    | String(x) -> String(x) |> conv_as_obj0
    | Timestamp(x) -> Timestamp(x)|> conv_as_obj0
    | Month(x) -> Month(x)|> conv_as_obj0
    | Date(x) -> Date(x)|> conv_as_obj0
    | DateTime(x) -> DateTime(x)|> conv_as_obj0
    | KTimespan(x) -> KTimespan(x)|> conv_as_obj0
    | Minute(x) -> Minute(x)|> conv_as_obj0
    | Second(x) -> Second(x)|> conv_as_obj0
    | TimeSpan(x) -> TimeSpan(x)|> conv_as_obj0
    | ABool(x) -> ABool(x) |> conv_as_array1 :> obj
    | AGuid(x) -> AGuid(x) |> conv_as_array1  :> obj
    | AByte(x) -> AByte(x) |> conv_as_array1  :> obj
    | AShort(x) -> AShort(x) |> conv_as_array1  :> obj
    | AInt(x) -> AInt(x) |> conv_as_array1  :> obj
    | ALong(x) -> ALong(x) |> conv_as_array1  :> obj
    | AReal(x) -> AReal(x) |> conv_as_array1  :> obj
    | AFloat(x) -> AFloat(x) |> conv_as_array1  :> obj
    | AChar(x) -> AChar(x) |> conv_as_array1  :> obj
    | AString(x) -> AString(x) |> conv_as_array1  :> obj
    | ATimestamp(x) -> ATimestamp(x) |> conv_as_array1  :> obj
    | AMonth(x) -> AMonth(x) |> conv_as_array1  :> obj
    | ADate(x) -> ADate(x) |> conv_as_array1  :> obj
    | ADateTime(x) -> ADateTime(x) |> conv_as_array1  :> obj
    | AKTimespan(x) -> AKTimespan(x) |> conv_as_array1  :> obj
    | AMinute(x) -> AMinute(x) |> conv_as_array1  :> obj
    | ASecond(x) -> ASecond(x) |> conv_as_array1  :> obj
    | ATimeSpan(x) -> ATimeSpan(x) |> conv_as_array1  :> obj
    | Dict(x,y) -> 
        let xc = conv_as_matrix0 x
        let yc = conv_as_matrix0 y
        let yc = match yc.Length with
                 | 1 -> yc
                 | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
        // let yc2 = Array2D.init r.[0].Length r.Length (fun i j -> r.[j].[i])
        // let r = yc |> Array.zip xc |> Array.map (fun (x,y) -> y |> Array.append x)
        let r = yc |> Array.append xc 
        // let m = r |> Array.map (fun x -> x.Length) |> Array.max
        let twoDimensionalArray = Array2D.init r.[0].Length r.Length   (fun i j -> r.[j].[i])
        twoDimensionalArray :> obj

    | Flip(x,y) -> 
        let xs = x |> Array.map (fun x -> [|x :> obj|]) 
        let yc = y |> List.map conv_as_array1  |> List.toArray
        // get the size an merge it together
        let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
        let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
        twoDimensionalArray :> obj        

    | AKObject(x) -> let x = x |> List.map conv_as_array1
                     let mn = x |> List.map (fun x -> x.Length) |> List.max
                     let r = x |> List.map (fun x -> [|for i in 1 .. (mn - x.Length) -> "" :> obj |] |> Array.append x) |> List.toArray
                     let twoDimensionalArray = Array2D.init r.Length r.[0].Length  (fun i j -> r.[i].[j])
                     twoDimensionalArray :> obj

