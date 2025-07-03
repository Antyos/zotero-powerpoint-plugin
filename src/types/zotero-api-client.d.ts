/**
 * TypeScript definitions for zotero-api-client
 *
 * This module provides a lightweight, minimalistic Zotero API client
 * developed in JavaScript with support for both Node and browser environments.
 *
 * @see https://github.com/tnajdek/zotero-api-client
 * @version 0.47.0
 */

// ==================== ZOTERO DATA TYPES ====================

declare module "zotero-api-client" {
  export interface ZoteroLink {
    href: string;
    type: string;
    attachmentType?: string;
    attachmentSize?: number;
  }

  export interface ZoteroLibrary {
    id: number;
    type: "user" | "group";
    name: string;
    links: Record<string, ZoteroLink>;
  }

  /**
   * Zotero Creator
   */
  export interface ZoteroCreator {
    creatorType: string;
    firstName?: string;
    lastName?: string;
    name?: string; // For single field names
  }

  /**
   * Zotero Tag
   */
  export interface ZoteroTag {
    tag: string;
    type?: number; // 0 for automatic, 1 for manual
    numItems?: number;
  }

  export interface ZoteroItemData {
    key: string;
    version: number;
    itemType: string;
    title: string;
    creators: ZoteroCreator[];
    abstractNote?: string;
    publicationTitle?: string;
    volume?: string;
    issue?: string;
    pages?: string;
    date?: string;
    series?: string;
    seriesTitle?: string;
    seriesText?: string;
    journalAbbreviation?: string;
    language?: string;
    DOI?: string;
    ISSN?: string;
    shortTitle?: string;
    url?: string;
    accessDate?: string;
    archive?: string;
    archiveLocation?: string;
    libraryCatalog?: string;
    callNumber?: string;
    rights?: string;
    extra?: string;
    tags: ZoteroTag[];
    collections: string[];
    relations: Record<string, string | string[]>;
    dateAdded: string;
    dateModified: string;
    [key: string]: any; // Allow for additional item type specific fields
  }

  /**
   * Zotero Item Type
   */
  export interface ZoteroItemRawResponse {
    version: number;
    key: string;
    data: Array<ZoteroItemData & { version: number }>;
    library: ZoteroLibrary;
    links: {
      alternate: ZoteroLink;
      attachment?: ZoteroLink;
      self?: ZoteroLink;
      [key: string]: ZoteroLink | undefined;
    };
    meta: {
      createdByUser?: {
        id: number;
        username: string;
        name: string;
      };
      creatorSummary: string;
      numChildren: number;
      parsedDate: string;
      [key: string]: unknown;
    };
  }

  /**
   * Zotero Collection
   */
  export interface ZoteroCollection {
    key: string;
    version: number;
    library: {
      type: "user" | "group";
      id: number;
      name: string;
      links: Record<string, any>;
    };
    links: Record<string, any>;
    meta: Record<string, any>;
    data: {
      key: string;
      version: number;
      name: string;
      parentCollection?: string;
      relations: Record<string, any>;
    };
  }

  /**
   * Zotero Group
   */
  export interface ZoteroGroup {
    id: number;
    version: number;
    links: Record<string, any>;
    meta: Record<string, any>;
    data: {
      id: number;
      version: number;
      name: string;
      description: string;
      url: string;
      hasImage: boolean;
      type: "Private" | "PublicOpen" | "PublicClosed";
      libraryEditing: "members" | "admins";
      libraryReading: "all" | "members";
      fileEditing: "none" | "members" | "admins";
      admins: Array<{ id: number; username: string; name: string }>;
      members: Array<{ id: number; username: string; name: string }>;
    };
  }

  /**
   * Zotero Search
   */
  export interface ZoteroSearch {
    key: string;
    version: number;
    library: {
      type: "user" | "group";
      id: number;
      name: string;
      links: Record<string, any>;
    };
    links: Record<string, any>;
    meta: Record<string, any>;
    data: {
      key: string;
      version: number;
      name: string;
      conditions: Array<{
        condition: string;
        operator: string;
        value: string;
      }>;
    };
  }

  /**
   * Item Type Information
   */
  export interface ItemType {
    itemType: string;
    localized: string;
  }

  /**
   * Item Field Information
   */
  export interface ItemField {
    field: string;
    localized: string;
  }

  /**
   * Creator Field Information
   */
  export interface CreatorField {
    field: string;
    localized: string;
  }

  /**
   * Item Type Field Information
   */
  export interface ItemTypeField {
    field: string;
    localized: string;
    baseField?: string;
  }

  /**
   * Creator Type Information
   */
  export interface CreatorType {
    creatorType: string;
    localized: string;
    primary?: boolean;
  }

  // ==================== API RESPONSE TYPES ====================

  /**
   * Base API Response
   */
  export interface ApiResponse {
    getResponseType(): string;
    getData(): any;
    getLinks(): any;
    getMeta(): any;
    getVersion(): number;
  }

  /**
   * Single Read Response
   */
  export interface SingleReadResponse<R = unknown, T = unknown> extends ApiResponse {
    options: RequestOptions;
    raw: R;
    response: Response;
    getResponseType(): "SingleReadResponse";
    getData(): T;
  }

  /**
   * Multi Read Response
   */
  export interface MultiReadResponse<R = unknown, T = unknown> extends ApiResponse {
    options: RequestOptions;
    raw: Array<R>;
    response: Response;
    getResponseType(): "MultiReadResponse";
    getData(): T[];
    getLinks(): any[];
    getMeta(): any[];
    getTotalResults(): string;
    getRelLinks(): Record<string, string>;
  }

  /**
   * Single Write Response
   */
  export interface SingleWriteResponse<T = unknown> extends ApiResponse {
    getResponseType(): "SingleWriteResponse";
    getData(): T;
  }

  /**
   * Multi Write Response
   */
  export interface MultiWriteResponse<T = unknown> extends ApiResponse {
    getResponseType(): "MultiWriteResponse";
    isSuccess(): boolean;
    getData(): T[];
    getLinks(): any;
    getMeta(): any;
    getErrors(): Record<string, any>;
    getEntityByKey(key: string): T;
    getEntityByIndex(index: number): T;
  }

  /**
   * Delete Response
   */
  export interface DeleteResponse extends ApiResponse {
    getResponseType(): "DeleteResponse";
  }

  /**
   * File Upload Response
   */
  export interface FileUploadResponse extends ApiResponse {
    getResponseType(): "FileUploadResponse";
    authResponse: any;
    response: any; // alias for authResponse
    uploadResponse: any;
    registerResponse: any;
    getVersion(): number;
  }

  /**
   * File Download Response
   */
  export interface FileDownloadResponse extends ApiResponse {
    getResponseType(): "FileDownloadResponse";
  }

  /**
   * File URL Response
   */
  export interface FileUrlResponse extends ApiResponse {
    getResponseType(): "FileUrlResponse";
  }

  /**
   * Raw API Response
   */
  export interface RawApiResponse extends ApiResponse {
    getResponseType(): "RawApiResponse";
  }

  /**
   * Pretend Response
   */
  export interface PretendResponse extends Omit<ApiResponse, "getVersion"> {
    getResponseType(): "PretendResponse";
    getVersion(): undefined;
  }

  /**
   * Error Response
   */
  export interface ErrorResponse extends Error {
    response: any;
    message: string;
    reason: string;
    options: any;
    getVersion(): number;
    getResponseType(): "ErrorResponse";
  }

  // ==================== API CONFIGURATION ====================

  /**
   * Request Options
   */
  export interface RequestOptions {
    // API Configuration
    apiScheme?: string;
    apiAuthorityPart?: string;
    apiPath?: string;
    authorization?: string;
    zoteroWriteToken?: string;
    ifModifiedSinceVersion?: string;
    ifUnmodifiedSinceVersion?: string;
    contentType?: string;

    // Query Parameters
    collectionKey?: string;
    content?: string;
    direction?: "asc" | "desc";
    format?:
      | "atom"
      | "bib"
      | "json"
      | "keys"
      | "versions"
      | "bibtex"
      | "bookmarks"
      | "coins"
      | "csljson"
      | "mods"
      | "refer"
      | "rdf_bibliontology"
      | "rdf_dc"
      | "rdf_zotero"
      | "ris"
      | "tei"
      | "wikipedia";
    include?: string;
    includeTrashed?: boolean;
    itemKey?: string;
    itemQ?: string;
    itemQMode?: "titleCreatorYear" | "everything";
    itemTag?: string | string[];
    itemType?: string;
    limit?: number;
    linkMode?: string;
    locale?: string;
    q?: string;
    qmode?: "titleCreatorYear" | "everything";
    searchKey?: string;
    since?: number;
    sort?:
      | "dateAdded"
      | "dateModified"
      | "title"
      | "creator"
      | "type"
      | "date"
      | "publisher"
      | "publicationTitle"
      | "journalAbbreviation"
      | "language"
      | "accessDate"
      | "libraryCatalog"
      | "callNumber"
      | "rights"
      | "addedBy"
      | "numItems";
    start?: number;
    style?: string;
    tag?: string | string[];

    // Special Options
    pretend?: boolean;
    uploadRegisterOnly?: boolean;
    retry?: number;
    retryDelay?: number;

    // Resource Configuration
    resource?: {
      top?: boolean;
      trash?: boolean;
      children?: boolean;
      groups?: boolean;
      itemTypes?: boolean;
      itemFields?: boolean;
      creatorFields?: boolean;
      itemTypeFields?: boolean;
      itemTypeCreatorTypes?: boolean;
      library?: boolean;
      collections?: boolean;
      items?: boolean;
      searches?: boolean;
      tags?: boolean;
      template?: boolean;
    };

    // Fetch Options
    method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
    body?: string;
    mode?: RequestMode;
    cache?: RequestCache;
    credentials?: RequestCredentials;
  }

  // ==================== API INTERFACE ====================

  /**
   * Main API Interface
   */
  export interface ZoteroApi<R = unknown, T = unknown> {
    // Configuration Methods
    (key?: string, opts?: RequestOptions): ZoteroApi;
    library(typeOrKey?: "user" | "group" | string, id?: number): ZoteroApi;

    // Resource Methods
    // If key is provided, returns a single item response, otherswise returns a multi-item response
    items(itemKey?: string): ZoteroApi<ZoteroItemRawResponse, ZoteroItemData>;
    collections(collectionKey?: string): ZoteroApi<ZoteroCollection>;
    subcollections(): ZoteroApi<ZoteroCollection>;
    searches(searchKey?: string): ZoteroApi<ZoteroSearch>;
    tags(tagName?: string): ZoteroApi<ZoteroTag>;
    settings(settingKey?: string): ZoteroApi;
    groups(): ZoteroApi<ZoteroGroup>;

    // Metadata Methods
    itemTypes(): ZoteroApi;
    itemFields(): ZoteroApi;
    creatorFields(): ZoteroApi;
    schema(): ZoteroApi;
    itemTypeFields(itemType: string): ZoteroApi;
    itemTypeCreatorTypes(itemType: string): ZoteroApi;
    template(itemType: string, subType?: string): ZoteroApi;

    // Modifier Methods
    top(): ZoteroApi;
    trash(): ZoteroApi;
    children(): ZoteroApi;
    publications(): ZoteroApi;
    deleted(): ZoteroApi;
    version(version: number): ZoteroApi;

    // File Methods
    attachment(
      fileName?: string,
      file?: ArrayBuffer,
      mtime?: number,
      md5sum?: string,
      patch?: ArrayBuffer,
      algorithm?: "xdelta" | "vcdiff" | "bsdiff"
    ): ZoteroApi;
    registerAttachment(
      fileName: string,
      fileSize: number,
      mtime: number,
      md5sum: string
    ): ZoteroApi;
    attachmentUrl(): ZoteroApi;

    // Utility Methods
    verifyKeyAccess(): ZoteroApi;
    getConfig(): RequestOptions;
    use(extend: (api: ZoteroApi) => ZoteroApi): ZoteroApi;

    // Execution Methods
    get(opts?: RequestOptions): Promise<SingleReadResponse<R, T> | MultiReadResponse<R, T>>;
    post(data: any[], opts?: RequestOptions): Promise<MultiWriteResponse<T>>;
    put(data: any, opts?: RequestOptions): Promise<SingleWriteResponse<T>>;
    patch(data: any, opts?: RequestOptions): Promise<SingleWriteResponse<T>>;
    del(keysToDelete?: string[], opts?: RequestOptions): Promise<DeleteResponse>;
    pretend(
      verb?: "get" | "post" | "put" | "patch" | "delete",
      data?: any,
      opts?: RequestOptions
    ): Promise<PretendResponse>;
  }

  // ==================== EXPORT ====================

  /**
   * Default export - the main API function
   */
  const api: ZoteroApi;
  export default api;

  /**
   * Request function (for advanced usage)
   */
  export function request(options: RequestOptions): Promise<ApiResponse>;
}
