// sp.ts
import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

// Update with your correct site URL and list name
const SiteURL = "https://postnord.sharepoint.com/sites/FutureTests";
const ListName = "Prompts";

/**
 * Represents a list item in the "Prompts" list.
 * Make sure your list has columns:
 * - Title (Single line of text)
 * - Description (Multiple lines of text)
 */
export interface FAQListItem {
  ID: number;        // Needed to identify the item for deletion
  Title: string;
  Description: string;
}

/**
 * Retrieve items (ID, Title, Description) from the "Prompts" list
 */
export const getFAQItems = async (sp: SPFI): Promise<FAQListItem[]> => {
  const web = Web([sp.web, SiteURL]);
  const items: FAQListItem[] = await web.lists
    .getByTitle(ListName)
    .items.select("ID", "Title", "Description")() as FAQListItem[];

  return items;
};

/**
 * Add a new item to the "Prompts" list
 */
export const addFAQItem = async (
  sp: SPFI,
  item: { Title: string; Description: string }
): Promise<void> => {
  const web = Web([sp.web, SiteURL]);
  await web.lists.getByTitle(ListName).items.add({
    Title: item.Title,
    Description: item.Description,
  });
};

/**
 * Delete an item from the "Prompts" list
 */
export const deleteFAQItem = async (sp: SPFI, itemId: number): Promise<void> => {
  const web = Web([sp.web, SiteURL]);
  await web.lists
    .getByTitle(ListName)
    .items.getById(itemId)
    .delete();
};
