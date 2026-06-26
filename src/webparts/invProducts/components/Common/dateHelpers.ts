export const formatDateIN = (dateStr: string) => {
  if (!dateStr) return "";

  const date = new Date(dateStr);

  //  timezone 
  const localDate = new Date(
    date.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })
  );

  const dd = ("0" + localDate.getDate()).slice(-2);
  const mm = ("0" + (localDate.getMonth() + 1)).slice(-2);
  const yyyy = localDate.getFullYear();

//   return `${dd}-${mm}-${yyyy}`;
  return `${dd}/${mm}/${yyyy}`; 
};