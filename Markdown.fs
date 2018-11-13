module Markdown

open System

type MarkdownContext = {
  chapter: int;
  headings: array<int>;
  listing: int;
}

let nextListing (context: MarkdownContext) = 
  let listing = context.listing
  let result = sprintf "%d-%d" context.chapter listing
  result, { context with listing = context.listing + 1}

let nextHeading (context: MarkdownContext, s: int) =
  context.headings.[s - 1] <- context.headings.[s - 1] + 1
  for i in s..(context.headings.Length - 1) do
    context.headings.[i] <- 0
  let rec genHeadingStr start acc =
    if start >= context.headings.Length || context.headings.[start] = 0
    then acc
    else genHeadingStr (start + 1) (sprintf "%s.%d" acc context.headings.[start])
  genHeadingStr 0 (context.chapter.ToString()), context

type MarkdownDocument = list<MarkdownBlock>
and MarkdownBlock =
  | Heading of int * MarkdownSpans
  | Paragraph of MarkdownSpans
  | ListBlock of list<MarkdownSpans>
  | CodeBlock of list<string> * string 
  | MetaBlock of list<string> * string
and MarkdownSpans = list<MarkdownSpan>
and MarkdownSpan =
  | Literal of string
  | InlineCode of string
  | Strong of MarkdownSpans
  | Emphasis of MarkdownSpans
  | HyperLink of MarkdownSpans * string
  | ImageLink of string * string

let (|StartsWith|_|) prefix input =
  let rec loop = function
    | p::prefix, r::rest when p = r ->
        loop (prefix, rest)
    | [], rest -> Some(rest)
    | _ -> None
  loop (prefix, input)

let rec parseBracketedBody closing acc = function
  | StartsWith closing (rest) ->
      Some(List.rev acc, rest)
  | c::chars ->
      parseBracketedBody closing (c::acc) chars
  | _ -> None

let parseBracketed opening closing = function
  | StartsWith opening chars ->
      parseBracketedBody closing [] chars
  | _ -> None

let (|Delimited|_|) delim = parseBracketed delim delim

let (|Image|_|) chars =
  match parseBracketed ['!'; '['] [']'] chars with
  | Some((alt, rest)) ->
    printfn "Bracketed: %A %A" alt rest
    match parseBracketed ['('] [')'] rest with
    | Some((link, chars)) -> Some((alt, link), chars)
    | _ -> None
  | _ -> None

let toString x = List.map string x |> String.concat ""

let rec parseSpans acc chars = seq {
  let emitLiteral = seq {
    if acc <> [] then
      yield acc |> List.rev |> toString |> Literal }
  
  match chars with
  | Delimited ['`'] (body, chars) ->
    yield! emitLiteral
    yield InlineCode(toString body)
    yield! parseSpans [] chars
  | Delimited ['*'; '*'] (body, chars)
  | Delimited ['_'; '_'] (body, chars) ->
    yield! emitLiteral
    yield Strong(parseSpans [] body |> List.ofSeq)
    yield! parseSpans [] chars
  | Delimited ['*'] (body, chars)
  | Delimited ['_'] (body, chars) ->
    yield! emitLiteral
    yield Emphasis(parseSpans [] body |> List.ofSeq)
    yield! parseSpans [] chars
  | Image ((alt, link), rest) ->
    yield ImageLink(toString alt, toString link)
    yield! parseSpans [] rest
  | c::chars ->
      yield! parseSpans (c::acc) chars
  | [] -> yield! emitLiteral
}

let partitionWhile f = 
  let rec loop acc = function
    | x::xs when f x -> loop (x::acc) xs
    | xs -> List.rev acc, xs
  loop []

let (|PrefixedLines|) (prefix:string) (lines: list<string>) =
  let prefixed, other = lines |> partitionWhile (fun line ->
    line.StartsWith(prefix))
  [ for line in prefixed ->
      line.Substring(prefix.Length) ], other

let (|LineSeparated|) lines =
  let isWhite = System.String.IsNullOrWhiteSpace
  match partitionWhile (isWhite >> not) lines with
  | par, _::rest
  | par, ([] as rest) -> par, rest

let (|AsCharList|) (str: string) = List.ofSeq str

let (|BoundedBy|_|) (start:string) (l: string) (lines: list<string>) = 
  match lines with
  | [] -> None
  | x::_ when not (x.StartsWith start) -> None
  | ss::xs -> 
    let (body, rest) = partitionWhile (fun (x: string) -> not (x.StartsWith l)) xs
    match rest with
    | [] -> None
    | x::_ when not (x.StartsWith l) -> None
    | x::rest -> Some((body, ss.Substring(start.Length), x, rest))

let parseHeader (lines: list<string>) =
  let body, rest = partitionWhile (fun (l: string) -> not (l.Trim().Length = 0)) lines
  let context: MarkdownContext = { chapter=0; headings=[|0; 0; 0; 0; 0; 0|]; listing=1 }
  let header = 
    List.fold (fun s (l:string) -> 
        let parts = l.Split('=')
        match parts.[0] with
        | "chapter" -> { s with chapter=int parts.[1] }
        | "heading1" -> 
          s.headings.[0] <- int parts.[1]
          s
        | "heading2" -> 
          s.headings.[1] <- int parts.[1]
          s
        | "heading3" -> 
          s.headings.[2] <- int parts.[1]
          s
        | "heading4" -> 
          s.headings.[3] <- int parts.[1]
          s
        | "heading5" -> 
          s.headings.[4] <- int parts.[1]
          s
        | "heading6" -> 
          s.headings.[5] <- int parts.[1]
          s
        | "listing" -> {s with listing=int parts.[1]}
        | _ -> s
      ) context body
  header, rest

let rec parseBlocks lines = seq {
  match lines with
  | AsCharList (StartsWith ['#'; ' '] heading)::lines ->
      yield Heading(1, parseSpans [] heading |> List.ofSeq)
      yield! parseBlocks lines
  | AsCharList(StartsWith ['#'; '#'; ' '] heading)::lines ->
      yield Heading(2, parseSpans [] heading |> List.ofSeq)
      yield! parseBlocks lines
  | PrefixedLines "    " (body, lines) when body <> [] ->
      yield CodeBlock(body, "")
      yield! parseBlocks lines
  | PrefixedLines "- " (body, lines) when body <> [] ->
      yield ListBlock(body |> List.map (List.ofSeq >> parseSpans [] >> List.ofSeq))
      yield! parseBlocks lines
  | BoundedBy "```" "```" (body, typ, _, rest) when body <> [] ->
      yield MetaBlock(body, typ)
      yield! parseBlocks rest
  | LineSeparated (body, lines) when body <> [] ->
      let body = String.concat " " body |> List.ofSeq
      yield Paragraph (parseSpans [] body |> List.ofSeq)
      yield! parseBlocks lines
  | line::lines when System.String.IsNullOrWhiteSpace(line) ->
      yield! parseBlocks lines
  | _ -> () 
}

let parseMarkdownDoc (lines: list<string>) = 
  let header, body = parseHeader lines
  header, parseBlocks body |> List.ofSeq
