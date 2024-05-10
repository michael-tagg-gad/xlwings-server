async function processResult(data) {
  await Office.onReady();
  try {
    return await Excel.run(async (context) => {
      let rawData = JSON.parse(data);
      console.log(rawData);

      // console.log(rawData);

      // Run Functions
      if (rawData !== null) {
        const forceSync = ["sheet"];
        for (let action of rawData["actions"]) {
          await globalThis.callbacks[action.func](context, action);
          if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
            await context.sync();
          }
        }
      }
    });
  } catch (error) {
    console.error(error);
  }
}
