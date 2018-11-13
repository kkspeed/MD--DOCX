module Doc

open System
open System.Drawing
open Markdown
open Novacode
open System.Xml.Linq
open System.Reflection

let w = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main")

let GetOrCreate_pPr (p: Paragraph) = 
  let pPr = p.Xml.Element( XName.Get( "pPr", w.NamespaceName ))
  if isNull pPr 
  then p.Xml.AddFirst( new XElement( XName.Get( "pPr", w.NamespaceName )))
       p.Xml.Element( XName.Get( "pPr", w.NamespaceName ) );
  else pPr

let applyTextProperty (p: Paragraph, name: XName, value: string, content: obj) =
  let method = p.GetType().GetMethod(
                 "ApplyTextFormattingProperty", 
                 BindingFlags.NonPublic ||| BindingFlags.Instance) 
  method.Invoke(p, [| name; value; content |]) |> ignore

let Shading (p: Paragraph, shading: Color) =
  let toHex (c: Color) = "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2")
  let pPr = GetOrCreate_pPr(p)
  let sh = pPr.Element( XName.Get( "shd", w.NamespaceName ) )
  let shd = if isNull sh
            then let shd = new XElement( XName.Get( "shd", w.NamespaceName ) )
                 pPr.Add shd
                 shd
            else sh
  let fillAttribute = shd.Attribute( XName.Get( "fill", w.NamespaceName ) )
  if isNull fillAttribute
  then shd.SetAttributeValue( XName.Get( "fill", w.NamespaceName ), toHex shading )
  else fillAttribute.SetValue( toHex shading )
  // applyTextProperty(p, XName.Get( "shd", w.NamespaceName ), "", 
  //     new XAttribute(XName.Get( "fill", w.NamespaceName ), toHex shading) ); 

let Numbered (p: Paragraph, id: int) =
  let numPr = new XElement(XName.Get("numPr", w.NamespaceName))
  let ilvl = new XElement(XName.Get("ilvl", w.NamespaceName))
  ilvl.SetAttributeValue(XName.Get("val", w.NamespaceName), "0")
  let numId = new XElement(XName.Get("numId", w.NamespaceName))
  numId.SetAttributeValue(XName.Get("val", w.NamespaceName), id.ToString())
  numPr.Add ilvl
  numPr.Add numId
  GetOrCreate_pPr(p).AddFirst numPr
  p.SpacingAfter(2.4) |> ignore
  p.SpacingBefore(2.4) |> ignore

let Indent (p: Paragraph, ind: int) =
  let pPr = GetOrCreate_pPr(p)
  let indent = new XElement(XName.Get("ind", w.NamespaceName))
  indent.SetAttributeValue(XName.Get("left", w.NamespaceName), ind.ToString())
  indent.SetAttributeValue(XName.Get("leftChars", w.NamespaceName), ind.ToString())
  pPr.Add indent

let writeSpan (p: Paragraph) = function
  | Literal text -> 
    p.Append(text) |> ignore
  | InlineCode code -> p.Append(code).Font("Courier New") |> ignore
  | ImageLink(alt, link) -> ()
  | Strong spans -> ()
  | Emphasis spans -> ()
  | HyperLink (spans, text) -> ()

// TODO: I don't know why for normal content, we'll need to set text style
// at each write..
let writeSpanNormal (p: Paragraph) (doc: DocX) = function
  | Literal text -> 
    p.Append(text).Font("SimSun").FontSize(10.5) |> ignore
  | InlineCode code -> p.Append(code).Font("Consolas") |> ignore
  | ImageLink(alt, link) -> 
    let image = doc.AddImage(link)
    Indent(p.AppendPicture(image.CreatePicture()), 0)
  | Strong spans -> ()
  | Emphasis spans -> ()
  | HyperLink (spans, text) -> ()

let rec writeBlock (context: MarkdownContext) (document: DocX) = function
  | Heading (level, spans) -> 
      let p = document.InsertParagraph ()
      p.Font("Microsoft YaHei") |> ignore
      match level with
        | 1 -> p.FontSize(26.0) |> ignore 
        | 2 -> p.FontSize(22.0) |> ignore
        | 3 -> p.FontSize(16.0) |> ignore
        | 4 -> p.FontSize(14.0) |> ignore
        | 5 -> p.FontSize(14.0) |> ignore
               p.Font("SimSun") |> ignore
        | _ -> p.FontSize(12.0) |> ignore
      let headerStr, newContext = nextHeading (context, level)
      p.StyleName <- sprintf "Heading%d" level
      p.Append(headerStr) |> ignore
      for span in spans do
        writeSpan p span
      newContext
  | Paragraph spans -> 
      let p1 = document.InsertParagraph ()
      for span in spans do
        writeSpanNormal p1 document span
      context
  | ListBlock spanList ->
      let id = spanList.GetHashCode()
      for spans in spanList do
        let p = document.InsertParagraph()
        Numbered(p, id)
        for s in spans do
          writeSpanNormal p document s
      context
  | MetaBlock (texts, meta) ->
      if meta = "提示"
      then
        let p = document.InsertParagraph()
        Shading(p.Append(meta).Font("SimSun").Color(Color.White)
          .FontSize(10.5).Bold(), Color.FromArgb(255, 227, 108, 10))
        p.StyleName <- "af6"
        for t in texts do
          let pp = document.InsertParagraph()
          pp.Font("Courier New").FontSize(9.0) |> ignore
          pp.StyleName <- "af5"
          Shading(pp, Color.FromArgb(255, 253, 233, 217))
          pp.Append(sprintf "%s" t).Font("SimSun").FontSize(10.5) |> ignore
        context
      else writeBlock context document (CodeBlock (texts, meta))
  | CodeBlock (texts, title) -> 
      let p = document.InsertParagraph()
      let codeTitle, newContext = nextListing context
      Shading(p.Append("代码清单" + codeTitle).Font("SimSun").Color(Color.White)
        .FontSize(10.5).Bold(), Color.FromArgb(255, 95, 73, 122))
      p.StyleName <- "af1"
      Indent(p, 0)
      let id = p.GetHashCode()
      for t in texts do
        let pp = document.InsertParagraph()
        pp.Font("Courier New").FontSize(9.0) |> ignore
        pp.StyleName <- "a3"
        Shading(pp, Color.FromArgb(255, 242, 242, 242))
        Numbered(pp, id)
        pp.Append(sprintf "%s" t).Font("Courier New").FontSize(9.0) |> ignore
      newContext