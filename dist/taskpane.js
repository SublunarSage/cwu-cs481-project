/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={27091:function(t){"use strict";t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},17991:function(t,e,r){"use strict";t.exports=r.p+"assets/Maggie.jpg"},60806:function(t,e,r){"use strict";t.exports=r.p+"8d768f65702f2137206f.css"}},e={};function r(n){var o=e[n];if(void 0!==o)return o.exports;var i=e[n]={exports:{}};return t[n](i,i.exports,r),i.exports}r.m=t,r.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return r.d(e,{a:e}),e},r.d=function(t,e){for(var n in e)r.o(e,n)&&!r.o(t,n)&&Object.defineProperty(t,n,{enumerable:!0,get:e[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),r.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;r.g.importScripts&&(t=r.g.location+"");var e=r.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&!t;)t=n[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=t}(),r.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){"use strict";e=function(){return n};var r,n={},o=Object.prototype,i=o.hasOwnProperty,a=Object.defineProperty||function(t,e,r){t[e]=r.value},c="function"==typeof Symbol?Symbol:{},u=c.iterator||"@@iterator",s=c.asyncIterator||"@@asyncIterator",f=c.toStringTag||"@@toStringTag";function l(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{l({},"")}catch(r){l=function(t,e,r){return t[e]=r}}function p(t,e,r,n){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),c=new P(n||[]);return a(i,"_invoke",{value:E(t,r,c)}),i}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}n.wrap=p;var d="suspendedStart",y="suspendedYield",g="executing",m="completed",v={};function w(){}function b(){}function x(){}var L={};l(L,u,(function(){return this}));var k=Object.getPrototypeOf,A=k&&k(k(I([])));A&&A!==o&&i.call(A,u)&&(L=A);var B=x.prototype=w.prototype=Object.create(L);function S(t){["next","throw","return"].forEach((function(e){l(t,e,(function(t){return this._invoke(e,t)}))}))}function T(e,r){function n(o,a,c,u){var s=h(e[o],e,a);if("throw"!==s.type){var f=s.arg,l=f.value;return l&&"object"==t(l)&&i.call(l,"__await")?r.resolve(l.__await).then((function(t){n("next",t,c,u)}),(function(t){n("throw",t,c,u)})):r.resolve(l).then((function(t){f.value=t,c(f)}),(function(t){return n("throw",t,c,u)}))}u(s.arg)}var o;a(this,"_invoke",{value:function(t,e){function i(){return new r((function(r,o){n(t,e,r,o)}))}return o=o?o.then(i,i):i()}})}function E(t,e,n){var o=d;return function(i,a){if(o===g)throw new Error("Generator is already running");if(o===m){if("throw"===i)throw a;return{value:r,done:!0}}for(n.method=i,n.arg=a;;){var c=n.delegate;if(c){var u=O(c,n);if(u){if(u===v)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===d)throw o=m,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=g;var s=h(t,e,n);if("normal"===s.type){if(o=n.done?m:y,s.arg===v)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(o=m,n.method="throw",n.arg=s.arg)}}}function O(t,e){var n=e.method,o=t.iterator[n];if(o===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=r,O(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),v;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,v;var a=i.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=r),e.delegate=null,v):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,v)}function N(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function W(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function P(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(N,this),this.reset(!0)}function I(e){if(e||""===e){var n=e[u];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,a=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=r,t.done=!0,t};return a.next=a}}throw new TypeError(t(e)+" is not iterable")}return b.prototype=x,a(B,"constructor",{value:x,configurable:!0}),a(x,"constructor",{value:b,configurable:!0}),b.displayName=l(x,f,"GeneratorFunction"),n.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===b||"GeneratorFunction"===(e.displayName||e.name))},n.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,l(t,f,"GeneratorFunction")),t.prototype=Object.create(B),t},n.awrap=function(t){return{__await:t}},S(T.prototype),l(T.prototype,s,(function(){return this})),n.AsyncIterator=T,n.async=function(t,e,r,o,i){void 0===i&&(i=Promise);var a=new T(p(t,e,r,o),i);return n.isGeneratorFunction(e)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},S(B),l(B,f,"Generator"),l(B,u,(function(){return this})),l(B,"toString",(function(){return"[object Generator]"})),n.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},n.values=I,P.prototype={constructor:P,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=r,this.done=!1,this.delegate=null,this.method="next",this.arg=r,this.tryEntries.forEach(W),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=r)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,o){return c.type="throw",c.arg=t,e.next=n,o&&(e.method="next",e.arg=r),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var a=this.tryEntries[o],c=a.completion;if("root"===a.tryLoc)return n("end");if(a.tryLoc<=this.prev){var u=i.call(a,"catchLoc"),s=i.call(a,"finallyLoc");if(u&&s){if(this.prev<a.catchLoc)return n(a.catchLoc,!0);if(this.prev<a.finallyLoc)return n(a.finallyLoc)}else if(u){if(this.prev<a.catchLoc)return n(a.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return n(a.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&i.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var a=o?o.completion:{};return a.type=t,a.arg=e,o?(this.method="next",this.next=o.finallyLoc,v):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),v},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),W(r),v}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;W(r)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:I(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=r),v}},n}function r(t,e,r,n,o,i,a){try{var c=t[i](a),u=c.value}catch(t){return void r(t)}c.done?e(u):Promise.resolve(u).then(n,o)}function n(t){return function(){var e=this,n=arguments;return new Promise((function(o,i){var a=t.apply(e,n);function c(t){r(a,o,i,c,u,"next",t)}function u(t){r(a,o,i,c,u,"throw",t)}c(void 0)}))}}function o(){return(o=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r.completed();case 1:case"end":return t.stop()}}),t)})))).apply(this,arguments)}function i(){return i=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection(),r.load(n),t.next=4,r.sync();case 4:return n.insertText("WARNING",Word.InsertLocation.end),n.font.color="red",t.next=8,r.sync();case 8:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),i.apply(this,arguments)}function a(){return c.apply(this,arguments)}function c(){return c=n(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r.document.body.insertParagraph("Test.",Word.InsertLocation.start),t.next=4,r.sync();case 4:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),c.apply(this,arguments)}function u(){return u=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i,a,c,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.body,(o=n.insertContentControl()).title="Attachment",o.tag="Figure",(i=n.insertParagraph("Figure, Figure Title",Word.InsertLocation.end)).alignment=Word.Alignment.center,(a=n.insertContentControl(i,"After")).title="Attach Image",a.tag="Image",a.appearance="Tags",a.insertHtml('<img src="" alt="Image Placeholder" style="width: 200px; height: auto; display: block; margin: 0 auto;">',Word.InsertLocation.end),i.insertBreak("Line","After"),(c=i.insertTable(1,1,"After")).getBorder(Word.BorderLocation.outside).type="Double",c.getBorder(Word.BorderLocation.all).width=1,(u=c.getRange("Start").insertParagraph("CAUTION","After")).alignment="Centered",u.font.bold=!0,u.getRange("After").insertText("Type here.","Start"),i.insertBreak("Line","After"),t.next=22,r.sync();case 22:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),u.apply(this,arguments)}function s(){return s=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection(),r.load(n),t.next=4,r.sync();case 4:return n.insertBreak("Line","After"),(o=n.insertTable(1,1,"After")).getBorder(Word.BorderLocation.outside).type="Single",o.getBorder(Word.BorderLocation.all).width=1,(i=o.getRange("Start").insertParagraph("NOTE","After")).alignment="Centered",i.font.bold=!0,i.getRange("After").insertText("Type here.","Start"),n.insertBreak("Line","After"),t.next=15,r.sync();case 15:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),s.apply(this,arguments)}function f(){return f=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection(),r.load(n),t.next=4,r.sync();case 4:return n.insertBreak("Line","After"),(o=n.insertTable(1,1,"After")).getBorder(Word.BorderLocation.outside).type="Double",o.getBorder(Word.BorderLocation.all).width=1,(i=o.getRange("Start").insertParagraph("CAUTION","After")).alignment="Centered",i.font.bold=!0,i.getRange("After").insertText("Type here.","Start"),n.insertBreak("Line","After"),t.next=15,r.sync();case 15:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),f.apply(this,arguments)}function l(){return l=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection(),r.load(n),t.next=4,r.sync();case 4:return n.insertBreak("Line","After"),(o=n.insertTable(1,1,"After")).getBorder(Word.BorderLocation.outside).type="Triple",o.getBorder(Word.BorderLocation.all).width=1,(i=o.getRange("Start").insertParagraph("WARNING","After")).alignment="Centered",i.font.bold=!0,i.getRange("After").insertText("Type here.","Start"),n.insertBreak("Line","After"),t.next=15,r.sync();case 15:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),l.apply(this,arguments)}function p(){return p=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection().paragraphs,r.load(n),t.next=4,r.sync();case 4:if(n.items[0].isListItem){t.next=9;break}return n.items[0].startNewList().load("$none"),t.next=9,r.sync();case 9:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),p.apply(this,arguments)}function h(){return h=n(e().mark((function t(r){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=r.document.getSelection().paragraphs,r.load(n),t.next=4,r.sync();case 4:if(n.items[0].isListItem){t.next=18;break}return(o=n.items[0].startNewList()).load("$none"),t.next=9,r.sync();case 9:return i=0,o.setLevelNumbering(0,Word.ListNumbering.arabic,[i,"."]),o.setLevelStartingNumber(0,1),i+=1,o.setLevelNumbering(1,"LowerLetter",[i,"."]),o.setLevelStartingNumber(1,1),o.load("levelTypes"),t.next=18,r.sync();case 18:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:r.completed();case 3:case"end":return t.stop()}}),t)}))),h.apply(this,arguments)}function d(){return y.apply(this,arguments)}function y(){return y=n(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=n(e().mark((function t(r){var n,o,i,a;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(n=r.document.sections).load("items"),t.next=4,r.sync();case 4:if(!(n.items.length>0)){t.next=20;break}return o=n.items[0],(i=o.getHeader("primary")).clear(),t.next=10,r.sync();case 10:return a=i.insertTable(1,3,"start",[["Procedure #","Procedure Title","Revision #\n"]]),t.next=13,r.sync();case 13:return a.font.bold=!0,a.getCell(0,1).horizontalAlignment="Centered",a.getCell(0,1).getBorder(Word.BorderLocation.right).type="None",a.getCell(0,1).getBorder(Word.BorderLocation.left).type="None",a.getCell(0,2).horizontalAlignment="Right",t.next=20,r.sync();case 20:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),y.apply(this,arguments)}function g(){Word.run((function(t){return t.document.body.insertHtml('\n      <div style="text-align: center; font-family: Arial; font-size: 12pt; line-height: 1;">\n          <p style="font-weight: bold; margin-bottom: 0.5em;">[Procedure Title]</p>\n          <p style="font-weight: bold; margin-bottom: 0.5em;">[Procedure Number]</p>\n          <p style="font-weight: bold; margin-bottom: 1.5em;">[Reactivity Statement]</p> <br>\n          <p style="margin-bottom: 0.5em;">Revision #</p>\n          <p style="margin-bottom: 0.5em;">[Safety or Quality Classification]</p>\n          <p style="margin-bottom: 1.5em;">Level of Use: </p> \n          <br><br><br>\n          \x3c!-- Additional Information (optional) --\x3e\n          <p style="margin-bottom: 0.5em;">Effective Date: </p>\n          <p style="margin-bottom: 0;">Responsible Organization: </p>\n          <p style="margin-bottom: 0.5em;">Prepared By: </p>\n          <p style="margin-bottom: 0.5em;">Approved By: </p>\n          <br>\n      </div>\n      ',"start"),t.sync()})).catch((function(t){console.log("Error: "+t.message),t instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(t.debugInfo))}))}Office.onReady((function(t){t.host===Office.HostType.Word&&(document.getElementById("test-btn").onclick=a,document.getElementById("header-btn").onclick=d,document.getElementById("cover-btn").onclick=g)})),Office.actions.associate("placeholder",(function(t){return o.apply(this,arguments)})),Office.actions.associate("test",(function(t){return i.apply(this,arguments)})),Office.actions.associate("insertAttachment",(function(t){return u.apply(this,arguments)})),Office.actions.associate("note",(function(t){return s.apply(this,arguments)})),Office.actions.associate("caution",(function(t){return f.apply(this,arguments)})),Office.actions.associate("warning",(function(t){return l.apply(this,arguments)})),Office.actions.associate("beginBullet",(function(t){return p.apply(this,arguments)})),Office.actions.associate("beginNumber",(function(t){return h.apply(this,arguments)}))}(),function(){"use strict";var t=r(27091),e=r.n(t),n=new URL(r(60806),r.b),o=new URL(r(17991),r.b);e()(n),e()(o)}()}();
//# sourceMappingURL=taskpane.js.map