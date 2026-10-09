const h = require('./form-selection-harness.cjs');
function arrayOwner(defaults, render, rootProps = {}, arrayProps = {}) {
  const F = h.suite('Array');
  let methods, array;
  function Owner() {
    methods = F.useForm({defaultValues: defaults});
    return h.React.createElement(F.Root, {methods, onSubmit: () => {}, ...rootProps},
      h.React.createElement(F.Array, {name: 'rows', ...arrayProps}, context => {array = context; return render?.(context, F);}));
  }
  const mounted = h.mount(h.React.createElement(Owner));
  return {...mounted, get methods() {return methods;}, get array() {return array;}};
}
module.exports = {arrayOwner};
