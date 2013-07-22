## <a name="pagetop" href="#pagetop">Arrays</a> ##

|Section				  				|Description							|
|:-------------------------------------:|---------------------------------------|
|[Fixed-Size Arrays](#fixed)			|An array with a fixed size				|
|[Dynamic Arrays](#dynamic)				|An array with a dynamic size			|
|[Single-Dimensional Arrays](#single)	|An array with a single dimension		|
|[Multidimensional Arrays](#multi)		|An array with more than one dimension	|
|[Jagged Arrays](#jagged)				|An array of arrays						|
----------

### <a name="fixed" href=#fixed>Fixed-Size Arrays</a> ###
An array with a size that cannot be changed at run-time.


```VB
Sub Example()
	'Create an array and set its size to 5
	'Because the array has been defined as an integer
	'Each item in the array must be an integer
    Dim a(0 To 4) As Integer
    
	'Assign a value to each point on the array
    a(0) = 1
    a(1) = 2
    a(2) = 3
    a(3) = 4
    a(4) = 5
    
    Debug.Print a(0)
End Sub
```

<sub>[Go to top](#pagetop)</sub>

----------


### <a name="single" href=#single>Single-Dimensional Arrays</a> ###

<sub>[Go to top](#pagetop)</sub>

----------