// Learn more about F# at http://fsharp.org

open System

open Markdown
open Doc
open Novacode
open System.IO

let testStr = @"
# Hello World
I am some text. This is paragraph 1.

This is paragraph 2.

    var a = new HelloWorld()
    printfn %A Yes No

Following

    var c = yes()
    function fs =
      | c = ss>>

Output

```scala
var y = yes()
var j = cool()
```

```提示
本部分内容很重要！
```

## Nope
"

[<EntryPoint>]
let main argv =
    if argv.Length < 3
    then printfn "Usage: dotnet run <original docx> <markdown> <output file>"
    else let template = argv.[0]
         let inputMd = argv.[1]
         let outputFile = argv.[2]
         use doc = DocX.Load(template)
         let context, block = 
           seq {
             use sr = new StreamReader(inputMd)
             while not sr.EndOfStream do
               yield sr.ReadLine()
           }
           |> List.ofSeq
           |> parseMarkdownDoc
         List.fold (fun c b ->
            writeBlock c doc b
         ) context block |> ignore
         doc.SaveAs(outputFile)
    0 
