/**
 * Image upload to the current site's "Site Assets" library, the same place
 * SharePoint itself keeps images added to pages.
 *
 * Only standard SharePoint REST calls are used, made as the signed-in author,
 * so the upload succeeds exactly when that person could add the file to Site
 * Assets by hand.
 */

import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import { uploadFileName, validateImageFile } from './imageFiles';

async function readJson<T>(response: SPHttpClientResponse, action: string): Promise<T> {
  if (!response.ok) {
    let detail = '';
    try {
      const body = await response.json();
      detail = (body && body.error && body.error.message) || '';
    } catch {
      // Body wasn't JSON; the status code is enough.
    }
    if (response.status === 401 || response.status === 403) {
      throw new Error(`You don't have permission to ${action} on this site. Ask a site owner, or paste a link to an image instead.`);
    }
    throw new Error(`Couldn't ${action} (error ${response.status}${detail ? `: ${detail}` : ''}).`);
  }
  return (await response.json()) as T;
}

/**
 * Uploads `file` to Site Assets and returns its server-relative URL
 * (e.g. "/sites/hr/SiteAssets/resource-card-20261030-logo.png").
 */
export async function uploadImageToSiteAssets(client: SPHttpClient, webUrl: string, file: File): Promise<string> {
  const problem = validateImageFile(file);
  if (problem) {
    throw new Error(problem);
  }
  const web = webUrl.replace(/\/+$/, '');

  // Creates the library if the site doesn't have one yet; returns it either way.
  const library = await readJson<{ Id: string }>(
    await client.post(`${web}/_api/web/lists/EnsureSiteAssetsLibrary`, SPHttpClient.configurations.v1, {}),
    'find the Site Assets library'
  );
  const folder = await readJson<{ ServerRelativeUrl: string }>(
    await client.get(`${web}/_api/web/lists(guid'${library.Id}')/RootFolder?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1),
    'find the Site Assets library'
  );

  const name = uploadFileName(file.name, file.type);
  const folderPath = encodeURIComponent(folder.ServerRelativeUrl.replace(/'/g, "''"));
  const uploaded = await readJson<{ ServerRelativeUrl: string }>(
    await client.post(
      `${web}/_api/web/GetFolderByServerRelativePath(decodedurl='${folderPath}')/Files/AddUsingPath(decodedurl='${encodeURIComponent(name)}',overwrite=false)`,
      SPHttpClient.configurations.v1,
      { body: file }
    ),
    'upload the image'
  );
  return uploaded.ServerRelativeUrl;
}
