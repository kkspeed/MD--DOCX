# MD -> docx Based on Supplied Template
You'll need DotNetCore and DocXCore in order to compile and run. 

```
dotnet add package DocXCore
```

Then run:
```
dotnet run <input.docx> <markdown> <output.docx>
```

It converts code to `af1` border with `a3` body.

Code blocks starts with 提示 are converted to `af6` header and `af5` body.
Headings are converted accordingly.

The markdown needs to come with a "header". The header needs to **finish with a
newline**. It could take keywords: `chapter, heading[1-6],listing`.

```
chapter=3
heading1=2
heading2=1
listing=7

```

The above header says this document is in Chapter 3. Currently in Section 3.2.1.
Code listing number starts at 7. So when appending the markdown content, it will
adjust numbering accordingly.

This example document:

~~~
chapter=3
heading1=2
heading2=1
listing=5

## First heading 2
content...

# Heading1
Paragraph 1

Paragraph 2

## Heading2

```提示
This is a hint.
```

```regular code
console.log("Hello World!")
console.log("Hello World!")
console.log("Hello World!")
```

- List item 1
- List item 2
- List item 3
~~~

will produce the following image:

![example](https://raw.githubusercontent.com/kkspeed/md--docx/master/image/sample1.png)

# Limitations
1. It's just a hack and not for genreic use!
2. No support for image, hyperlink, cross ref etc.