args:.Q.def[`name`port!("name";8866);].Q.opt .z.x

/ remove this line when using in production
/ name:localhost:8866::
{ if[not x=0; @[x;"\\\\";()]]; value"\\p 8866"; } @[hopen;`:localhost:8866;0];

\l bt.q
r) library(ggplot2)

cons:flip `address`userid`handle`arg!()

.z.po:{0N!"Port opened\n";0N!r:(.z.a;.z.u;.z.w;x);`cons insert r;}
.z.pc:{0N!"Port closed\n";0N!r:(.z.a;.z.u;.z.w;x);delete from `cons where handle = x;x}
.z.ps:{[x]0N!(`zps;x);value x}
.z.pg:{[x]0N!(`zpg;x);value x}

.u.sub:{[x;y] .u.w:.z.w;.u.x:x;.u.y:y}

/ (neg .u.w) ("upd";`abc;([]a:1 2 3;b:4 5 6))

.bt.outputTrace:.bt.outputTrace1

sym:`abc`acb`cab`bca

(::)atoms:update t:i from ([] c: .Q.t)

(::)atoms:select from atoms where not null c 

(::)atoms:update name:key each c$\:() from atoms

(::)atoms:update ndata:{first 1#x}each c$\:() from atoms

(::)atoms:update data:{first .[?;(1;x);1#x]}@'ndata from atoms

(::)atoms:update data:(1h;1i;1j;1e;1f) from atoms where name in `short`int`long`real`float

(::)atoms:update data:(.z.p;`month$.z.p;`date$.z.p;`datetime$.z.p;`timespan$.z.p;`minute$.z.p;`second$.z.p;`time$.z.p;`sym?`abc) from atoms where t in 12 13 14 15 16 17 18 19 20

(::)atoms:update ldata:3#/:data from atoms