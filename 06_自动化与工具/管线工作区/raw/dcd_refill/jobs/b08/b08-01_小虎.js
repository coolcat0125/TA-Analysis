// DCD 抽取任务 · 小虎
// 通道：Tabbit 浏览器（配方见 Update/DCD参数回填_进度跟踪.md）
// 用法：tabbit-cli.exe nodejs --task dcd-params --request-id b08-01 --timeout-ms 180000 < 本文件
(async () => {
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const out = { kw: "小虎", series_url: null, cars: [] };
  try {
    // 1) 搜索页取 series_id
    const searchUrl = "https://www.dongchedi.com/search?keyword=" + encodeURIComponent("小虎");
    // （浏览器上下文内执行：导航→等待→取 __NEXT_DATA__ 或 DOM 中的 series 链接）
    // 2) 参数页取 rawData
    //    document.querySelectorAll('script') 找 __NEXT_DATA__ → props.pageProps.rawData
    //    rawData.car_info[] → info[key].value；rawData.properties[] → 车型目录
    // 3) 款型遍历：cars.push({ id, name, ff, vals })
    // 期望输出：JSON（stdout）— 由 Tabbit 回传后落盘 raw/dcd_refill/dcd_s<sid>.json
    console.log(JSON.stringify(out));
  } catch (e) {
    console.log(JSON.stringify({ kw: "小虎", error: String(e) }));
  }
})();
