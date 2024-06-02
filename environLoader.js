// eslint-disable-next-line @typescript-eslint/no-var-requires
const { envDefault, envLocal, envAzure } = require('./privateEnvOptions.js');

module.exports = function(source) {
  // @ts-ignore
  const qstr = this.query || 'not defined';
  const match = qstr.match(/app_deploy=(.*)/i);
  let deployValue = 'localhost';
  if (match) {
    deployValue = match[1].trim();
  } 

  if (deployValue === 'localhost' ) {
    return source.toString()
      .replace(new RegExp(envDefault.ClientId, 'g'), envLocal.ClientId);
  } else {
    return source.toString()
      .replace(new RegExp(envDefault.ClientId, 'g'), envAzure.ClientId);
  }
};