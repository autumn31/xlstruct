# xlstruct
xlstruct is a package that uses [xlsx (credit to tealeg)](github.com/tealeg/xlsx) to deal with unmarshaling xlsx sheet.  
By specifying a row as schema, it will unmarshal each row to stuct and returns slice of the unmarshaled structs.  

### example  
if we have a sheet which is:
|id|name|age|
|-|-|-|
|1|John|10|
|2|Eason|20|
|3|Sandy|12|

you can get the slice of students easily by defining the struct
```go
type Student struct{
  ID    string  `excel:"id"`
  Name  string  `excel:"name"`
  Age   int     `excel:"age"`
}
```
and then

```go
rs := []Student{}
err := xlstruct.Unmarshal(&rs, sh, 0) // 0 => row 0 is header row, sh is sheet
if err != nil {
  panic(err)
}

fmt.Println(rs)
```
