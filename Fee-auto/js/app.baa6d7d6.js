(function(t){function e(e){for(var n,u,i=e[0],c=e[1],s=e[2],l=0,f=[];l<i.length;l++)u=i[l],Object.prototype.hasOwnProperty.call(r,u)&&r[u]&&f.push(r[u][0]),r[u]=0;for(n in c)Object.prototype.hasOwnProperty.call(c,n)&&(t[n]=c[n]);p&&p(e);while(f.length)f.shift()();return o.push.apply(o,s||[]),a()}function a(){for(var t,e=0;e<o.length;e++){for(var a=o[e],n=!0,i=1;i<a.length;i++){var c=a[i];0!==r[c]&&(n=!1)}n&&(o.splice(e--,1),t=u(u.s=a[0]))}return t}var n={},r={app:0},o=[];function u(e){if(n[e])return n[e].exports;var a=n[e]={i:e,l:!1,exports:{}};return t[e].call(a.exports,a,a.exports,u),a.l=!0,a.exports}u.m=t,u.c=n,u.d=function(t,e,a){u.o(t,e)||Object.defineProperty(t,e,{enumerable:!0,get:a})},u.r=function(t){"undefined"!==typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},u.t=function(t,e){if(1&e&&(t=u(t)),8&e)return t;if(4&e&&"object"===typeof t&&t&&t.__esModule)return t;var a=Object.create(null);if(u.r(a),Object.defineProperty(a,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var n in t)u.d(a,n,function(e){return t[e]}.bind(null,n));return a},u.n=function(t){var e=t&&t.__esModule?function(){return t["default"]}:function(){return t};return u.d(e,"a",e),e},u.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},u.p="";var i=window["webpackJsonp"]=window["webpackJsonp"]||[],c=i.push.bind(i);i.push=e,i=i.slice();for(var s=0;s<i.length;s++)e(i[s]);var p=c;o.push([0,"chunk-vendors"]),a()})({0:function(t,e,a){t.exports=a("56d7")},"56d7":function(t,e,a){"use strict";a.r(e);var n=a("1da1"),r=(a("a52c"),a("2ed4")),o=(a("537a"),a("ac28")),u=(a("5246"),a("6b41")),i=(a("66b9"),a("b650")),c=(a("e260"),a("e6cf"),a("cca6"),a("a79d"),a("96cf"),a("e9c4"),a("d3b7"),a("2b0e")),s=function(){var t=this,e=t.$createElement,a=t._self._c||e;return a("div",{attrs:{id:"app"}},[a("router-view")],1)},p=[],l={name:"App"},f=l,v=a("2877"),b=Object(v["a"])(f,s,p,!1,null,null,null),m=b.exports,d=a("8c4f"),h=function(){var t=this,e=t.$createElement,a=t._self._c||e;return a("div",{staticClass:"app-page"},[a("router-view"),a("van-tabbar",{attrs:{route:""}},[a("van-tabbar-item",{attrs:{replace:"",to:"/home",icon:"home-o"}},[t._v("首页")]),a("van-tabbar-item",{attrs:{replace:"",to:"/timer",icon:"todo-list-o"}},[t._v("任务")]),a("van-tabbar-item",{attrs:{replace:"",to:"/user",icon:"user-o"}},[t._v("我的")])],1)],1)},y=[],g={name:"Index"},_=g,w=Object(v["a"])(_,h,y,!1,null,"35026932",null),O=w.exports,x=function(){var t=this,e=t.$createElement,a=t._self._c||e;return a("div",{staticClass:"app-page"},[a("van-nav-bar",{attrs:{title:"首页"}}),a("van-button",{attrs:{type:"primary"},on:{click:t.handleTest}},[t._v("淘宝5")])],1)},$=[],j={name:"Home",methods:{handleTest:function(){var t="淘宝",e="com.taobao.taobao";this.$auto('app.launchPackage("'.concat(e,'")'))?(this.$auto('toast("启动'.concat(t,'中")')),this.$waitFor('className("android.widget.TabHost")',3e4),this.$auto('toast("开始，按【音量减】键可中止")')):this.$auto('toast("未找到'.concat(t,'应用，请安装后再试")'))}}},k=j,P=Object(v["a"])(k,x,$,!1,null,"45766fd3",null),T=P.exports,E=function(){var t=this,e=t.$createElement,a=t._self._c||e;return a("div",{staticClass:"app-page"},[a("van-nav-bar",{attrs:{title:"定时任务"}}),a("van-button",{attrs:{type:"primary"}},[t._v("主要按钮")])],1)},S=[],C={name:"Timer"},M=C,J=Object(v["a"])(M,E,S,!1,null,"7ebe5326",null),F=J.exports,H=function(){var t=this,e=t.$createElement,a=t._self._c||e;return a("div",{staticClass:"app-page"},[a("van-nav-bar",{attrs:{title:"我的"}}),a("van-button",{attrs:{type:"primary"}},[t._v("主要按钮")])],1)},N=[],R={name:"User"},A=R,I=Object(v["a"])(A,H,N,!1,null,"5f340731",null),U=I.exports;c["a"].use(d["a"]);var q=new d["a"]({base:"",routes:[{path:"/",name:"index",component:O,children:[{path:"/home",name:"home",component:T},{path:"/timer",name:"timer",component:F},{path:"/user",name:"user",component:U}]}]}),z=(a("fb98"),void 0);c["a"].use(i["a"]),c["a"].use(u["a"]),c["a"].use(o["a"]),c["a"].use(r["a"]),c["a"].config.productionTip=!1,c["a"].prototype.$auto=function(t,e){var a=void 0;try{a=prompt(t,JSON.stringify(e))}catch(n){alert(n)}return a},c["a"].prototype.$sleep=function(t){return new Promise((function(e){setTimeout((function(){e(!0)}),t)}))},c["a"].prototype.$waitFor=function(){var t=Object(n["a"])(regeneratorRuntime.mark((function t(e){var a,n,r,o=arguments;return regeneratorRuntime.wrap((function(t){while(1)switch(t.prev=t.next){case 0:a=o.length>1&&void 0!==o[1]?o[1]:1e4,n=0,r=!1;case 3:if(!z.$auto("".concat(e,".exists()"))){t.next=9;break}return r=!0,n=a,t.abrupt("break",12);case 9:z.$sleep(500),n+=500;case 11:if(n<a){t.next=3;break}case 12:return t.abrupt("return",r);case 13:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}(),q.beforeEach((function(t,e,a){"/"===t.path?a({path:"/home"}):a()})),new c["a"]({router:q,render:function(t){return t(m)}}).$mount("#app")},fb98:function(t,e,a){}});