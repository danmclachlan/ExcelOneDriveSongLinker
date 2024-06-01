const { envLocal, envAzure } = require('./privateEnvOptions.js');

module.exports = function(source) {
  const logger = this.getLogger('environLoader') || console;

  const qstr = this.query || 'not defined';
  const match = qstr.match(/app_deploy=(.*)/i);
  let deployValue = 'localhost';
  if (match) {
    deployValue = match[1].trim();
  } 

  //logger.info(`platform: ${platformValue}`);

  if (deployValue === 'localhost' ) {
    return source;
  } else {
    return source.toString()
      .replace(new RegExp(envLocal.ClientId, 'g'), envAzure.ClientId)
      .replace(new RegExp(envLocal.Url, 'g'), envAzure.Url);
  }
};