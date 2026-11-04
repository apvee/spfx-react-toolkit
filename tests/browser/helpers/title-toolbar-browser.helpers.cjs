const settle = page => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function settlePalette(page) {
  await page.locator('[data-title-color-picker]').waitFor();
  await page.waitForFunction(() => {
    for (let node = document.querySelector('[data-title-color-picker]'); node; node = node.parentElement)
      if (node.getAnimations().some(animation => animation.playState === 'running')) return false;
    return true;
  });
}
module.exports = { settle, settlePalette };
