export function getQueryParams(search) {
  let params = {};
  if (search) {
    search = search.substring(1); // 先頭の?を除去

    search.split("&").forEach(param => {
      const s = param.split("=");
      params = {
        ...params,
        [s[0]]: s[1]
      };
    });
  }
  console.log("params");
  console.log(params);
  return params;
}
