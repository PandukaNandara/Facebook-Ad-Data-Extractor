const minifyJson = (inJson = '') => {
  return inJson.replace(/[\r\n]+/g," ");
};

export default minifyJson;
