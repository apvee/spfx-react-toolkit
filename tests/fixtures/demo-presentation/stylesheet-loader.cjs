// Compile the app's actual Sass rather than replacing its module with an empty object.
const sass = require('sass');
module.exports = function(source) {
  const css = sass.renderSync({ data: source }).css.toString();
  const classes = Object.fromEntries([...css.matchAll(/\.([\w-]+)/g)].map(match => [match[1], match[1]]));
  return `const style = document.createElement('style'); style.textContent = ${JSON.stringify(css)}; document.head.appendChild(style); export default ${JSON.stringify(classes)};`;
};
