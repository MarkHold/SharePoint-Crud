import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

export interface FAQListItem {
  ID: string;
  Title: string;
  Category: string;
  // change to Additional_x0020_Contact_x0028_s when publishing
  Additional_x0020_Contact: {
    Title: string;
    ID: string;
    EMail: string;
  };
  Listplace: string;
  ITSM_x0020_number: string;
  To_x0020_Date: string;
  From_x0020_Date: string;
  Description: string;
  Targetgroup: string[] | undefined;
}

const SiteURL = "https://postnord.sharepoint.com/sites/pn-broadcast";
const ListName = "NSDTasks";

export const getFAQItems = async (sp: SPFI) => {
  // Create a Web instance combining sp.web with the Site URL
  const web = Web([sp.web, SiteURL]);

  const now = new Date().toISOString();

  // Utility function to format dates as "YYYY-MM-DD HH:MM"
  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);

    const year = date.getFullYear();
    const month = (date.getMonth() + 1 < 10 ? "0" : "") + (date.getMonth() + 1);
    const day = (date.getDate() < 10 ? "0" : "") + date.getDate();
    const hours = (date.getHours() < 10 ? "0" : "") + date.getHours();
    const minutes = (date.getMinutes() < 10 ? "0" : "") + date.getMinutes();

    return `${year}-${month}-${day} ${hours}:${minutes}`;
  };

  // Retrieve the FAQ items from the list using a filter on To_x0020_Date and Listplace
  const items: FAQListItem[] = await web.lists
    .getByTitle(ListName)
    .items
    .filter(`To_x0020_Date ge datetime'${now}' and Listplace eq 'Open'`)
    .select(
      "ID",
      "Title",
      "Category",
      "Description",
      "Targetgroup",
      "Additional_x0020_Contact",
      "Additional_x0020_Contact/EMail",
      "ITSM_x0020_number",
      "To_x0020_Date",
      "From_x0020_Date",
      "Listplace"
    )
    .expand("Additional_x0020_Contact")();

  console.log(items);

  return items.map((item) => {
    return {
      ...item,
      From_x0020_Date: formatDate(item.From_x0020_Date),
      To_x0020_Date: formatDate(item.To_x0020_Date),
      Targetgroup: item.Targetgroup?.map((groupname) => {
        return groupname.toLocaleLowerCase();
      }),
    };
  });
};
