import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFx, spfi } from "@pnp/sp";
import { ISharePointItem } from "../BulkUploadSpFx";

export const getItemsFromList = async (
  context: WebPartContext,
  listName: string
): Promise<ISharePointItem[]> => {
  const sp = spfi().using(SPFx(context));
  const query = sp.web.lists
    .getByTitle(listName)
    .items.top(999)
    .orderBy("ID", false);
  let allItems: ISharePointItem[] = [];

  // Process all pages
  for await (const page of query) {
    // 'page' is an array of items from the current batch
    allItems = allItems.concat(page);
  }

  // console.log("Total items retrieved:", allItems.length);
  return allItems;
};

export const saveData = async (
  data: ISharePointItem | any,
  context: WebPartContext,
  listName: string
) => {
  const sp = spfi().using(SPFx(context));
  const result = await sp.web.lists.getByTitle(listName).items.add({
    Title: data.Title,
    FirstName: data.FirstName,
    LastName: data.LastName,
    WorkEmail: data.WorkEmail,
    PersonalEmail: data.PersonalEmail,
    BirthDate: data.BirthDate,
    HireDate: data.HireDate,
    WorkMode: data.WorkMode,
    Salary: data.Salary,
    IsMarried: data.IsMarried,
    SocialProfile: {
      Url: data.SocialProfile,
    },
    JobTitle: data.JobTitle,
    About: data.About,
  });

  return result; // result.data contains the created item info
};
