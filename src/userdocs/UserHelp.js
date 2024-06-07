// @ts-ignore
import * as generatedHtml from './UserHelp.md';
let div = document.createElement('div');
div.innerHTML = generatedHtml['default'];
document.body.appendChild(div);
document.title = 'Excel OneDrive Song Linker';