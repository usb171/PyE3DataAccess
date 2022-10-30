# PyE3DataAccess, manipulando a biblioteca de comunicação E3DataAccess com Python

## Requisitos 
```
pip install comtypes 
```

## Método ReadValue

```
if __name__ == '__main__':
	pyE3DataAccess = PyE3DataAccess(server="localhost")
	print(pyE3DataAccess.readvalue(pathname="SUBESTACAO.B11C1.[11C1].Terminal1.IA"))
```

## Método WriteValue

```
if __name__ == '__main__':
	pyE3DataAccess = PyE3DataAccess(server="localhost")
	tag = "SUBESTACAO.B11C1.[31C1-4].Measurements.PosicaoChave.Operator"
	date = time.strftime("%d-%m-%Y %H:%M:%S", time.gmtime())
	pyE3DataAccess.writevalue(pathname=tag, date=date, quality=192, value=1)
```

## Método ExecuteQuery

```
if __name__ == '__main__':
    pyE3DataAccess = PyE3DataAccess(server="localhost")
    consulta = "Dados.Consulta1"
    result = pyE3DataAccess.executeQuery(pathname=consulta)
    result = [result[0][0],result[0][1]]
    for row in range(len(result[0])):
        print(result[0][row], result[1][row])
```

## Métodos RegisterCallBack e UnregisterCallBack

```
  if __name__ == '__main__':
    pyE3DataAccess = PyE3DataAccess(server="localhost")
    pyE3DataAccess.registerCallback(pathname="SUBESTACAO.B11C1.[11C1].Terminal1.IA")
    pyE3DataAccess.pumpEventsChanged()
```





