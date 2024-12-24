import { MSGraphClientV3 } from "@microsoft/sp-http";

interface Graphresult {
  value: [
    {
      id: string;
      displayName: string;
    }
  ];
}

export const getCurrentUserGroups = async (client: MSGraphClientV3) => {
  const data: Graphresult = await client
    .api('me/memberOf?$search="displayName:PN"')
    .header("ConsistencyLevel", "eventual")
    .select("id, displayName")
    .get();

  return data;
};

//const msgraphclient = await this.context.msGraphClientFactory.getClient(
//"3"
//);
