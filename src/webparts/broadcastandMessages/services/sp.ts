import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

export interface FAQListItem {
  ID: string;
  Title: string;
  Status: string;
  ITSMnumber: string;
  To_x0020_Date: string;
  From_x0020_Date: string;
  Description: string;
  Targetgroup: string[] | undefined;
  ListSource: string; // <-- New property
}

const SiteURL =
  "https://postnord.sharepoint.com/sites/pn-broadcast-testenvironment";

const LISTS = {
  Changes: "Changes",
  ServiceMessages: "Service Messages",
};

export const getFAQItems = async (sp: SPFI) => {
  const web = Web([sp.web, SiteURL]);

  // Example offset of 2 days
  const today = new Date();
  today.setDate(today.getDate() - 2);
  const now = today.toISOString();

  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const year = date.getFullYear();
    const month = (date.getMonth() + 1 < 10 ? "0" : "") + (date.getMonth() + 1);
    const day = (date.getDate() < 10 ? "0" : "") + date.getDate();
    const hours = (date.getHours() < 10 ? "0" : "") + date.getHours();
    const minutes = (date.getMinutes() < 10 ? "0" : "") + date.getMinutes();
    return `${year}-${month}-${day} ${hours}:${minutes}`;
  };

  const filterQuery = `To_x0020_Date ge datetime'${now}' and Status eq 'Open'`;

  // Fetch Changes items
  const changesItems: FAQListItem[] = await web.lists
    .getByTitle(LISTS.Changes)
    .items.filter(filterQuery)
    .select(
      "ID",
      "Title",
      "Description",
      "Targetgroup",
      "ITSMnumber",
      "To_x0020_Date",
      "From_x0020_Date",
      "Status"
    )();

  // Fetch Service Messages items
  const serviceMessagesItems: FAQListItem[] = await web.lists
    .getByTitle(LISTS.ServiceMessages)
    .items.filter(filterQuery)
    .select(
      "ID",
      "Title",
      "Description",
      "Targetgroup",
      "ITSMnumber",
      "To_x0020_Date",
      "From_x0020_Date",
      "Status"
    )();

  // Tag each item with its source list
  const taggedChangesItems = changesItems.map((item) => ({
    ...item,
    ListSource: LISTS.Changes,
    From_x0020_Date: formatDate(item.From_x0020_Date),
    To_x0020_Date: formatDate(item.To_x0020_Date),
    Targetgroup: item.Targetgroup?.map((g) => g.toLocaleLowerCase()),
  }));

  const taggedServiceMessagesItems = serviceMessagesItems.map((item) => ({
    ...item,
    ListSource: LISTS.ServiceMessages,
    From_x0020_Date: formatDate(item.From_x0020_Date),
    To_x0020_Date: formatDate(item.To_x0020_Date),
    Targetgroup: item.Targetgroup?.map((g) => g.toLocaleLowerCase()),
  }));

  // Combine them
  const combinedItems = [...taggedChangesItems, ...taggedServiceMessagesItems];
  return combinedItems;
};
