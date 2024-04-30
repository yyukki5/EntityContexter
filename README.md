# Entity Contexter
I'd like to create VBA class from database, Automatically. (it's Scaffolding)  
To be implement DbContext.

## Sample Code
~~~
dim ec As new EntityContexter
ec.Init(connectionString).ScaffoldDataBese
~~~
Then created database table's classes. 


## Japanese Note
VBA でも EF の Scaffolding をしたい。という気持ちのクラス。  
Scaffolding 出来たので一旦作業を終了。  
DbContext はまだない。いつか実装できればいいかも。  
