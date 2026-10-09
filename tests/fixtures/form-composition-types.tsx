import * as React from 'react';
import {createForm} from '../../packages/spfx-react-toolkit/src/controls/form';
const Order=createForm<{title:string; created:Date; color:{h:number;s:number;v:number}; tags:string[]; customer?:{name:string;lines:{id:string;quantity:number;detail:{sku:string}}[]};rows:{name:string;active:boolean}[]}>();
<Order.Object name="customer">{({Fields})=><>
  <Fields.Input name="name" />
  <Fields.Array name="lines">{({items,append})=>{append({id:'a',quantity:1,detail:{sku:'a'}});return items.map(({key,Fields})=><Fields.Object key={key} name="detail">{({Fields})=><Fields.Input name="sku"/>}</Fields.Object>);}}</Fields.Array>
  {/* @ts-expect-error Root names are not relative object paths. */}
  <Fields.Input name="title" />
</>}</Order.Object>;
<Order.Array name="rows">{({items,append,update,replace})=>{append({name:'a',active:true}); update(0,{name:'b',active:false});replace([]);
  // @ts-expect-error Whole item model must be supplied.
  append({name:'missing active'});
  return items.map(({key,Fields})=><React.Fragment key={key}><Fields.Input name="name"/><Fields.Checkbox name="active"/>
    {/* @ts-expect-error Boolean field cannot use a string input. */}
    <Fields.Input name="active"/>
  </React.Fragment>);
}}</Order.Array>;
// @ts-expect-error Scalars are not object scopes.
<Order.Object name="title">{()=>undefined}</Order.Object>;
// @ts-expect-error Native Date is not an object record scope.
<Order.Object name="created">{()=>undefined}</Order.Object>;
// @ts-expect-error HSV is a scalar color domain, not an object scope.
<Order.Object name="color">{()=>undefined}</Order.Object>;
// @ts-expect-error Primitive arrays belong to multi-value controls.
<Order.Array name="tags">{()=>undefined}</Order.Array>;
// @ts-expect-error Array scope requires an array path.
<Order.Array name="customer">{()=>undefined}</Order.Array>;
