export async function getActiveBookName() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load("name");
      await context.sync();
      return workbook.name;
    });
  } catch (error) {
    console.error(error);
  }
}
