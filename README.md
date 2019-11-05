convert csv to xlsx(Microsoft Office Excel)

```
Usage of csv2xlsx:
  -c, --column []string   set column type
                          i:int, f:float, s:string
                          example: -cA:i -c2:s --column=C:f
      --header string     set header
                          0:no header 1:has header
                           (default "1")
  -o, --output string     output dir (default "./")
```

example.csv

```csv
id,title,price,created_at
1,mouse,299.90,2006-01-02
2,keyboard,120,2006-01-03
3,502,3,2006-01-04
```

convert one file:

```
csv2xlsx example.csv
```

convert some files:

```
csv2xlsx --output=/tmp --column=A:i --column=B:s -cC:f --header=1 \ 
example.csv example2.csv example3.csv
```
