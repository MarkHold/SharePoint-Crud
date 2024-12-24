// spAddItem.ts
import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

const SiteURL = "https://postnord.sharepoint.com/sites/FutureTests";
const PromptsListName = "Prompts";

export interface PromptsListItem {
  Title: string;
  Description: string;
}

export const addPromptsItem = async (
  sp: SPFI,
  item: PromptsListItem
): Promise<void> => {
  const web = Web([sp.web, SiteURL]);
  await web.lists
    .getByTitle(PromptsListName)
    .items.add({
      Title: item.Title,
      Description: item.Description,
    });
};
