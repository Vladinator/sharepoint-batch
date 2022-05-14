export type CallbackOptions = {
    before?: Function;
    done?: Function;
    fail?: Function;
    finally?: Function;
}

export type CallbackProps = 'before' | 'done' | 'fail' | 'finally';

export type RequestOptions = CallbackOptions & RequestInit & {
    method: 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';
    url: string;
}

export type RequestResponse = Response | undefined;

export type SharePointOptions = {
    url: string;
    digest: string;
};

export type SharePointParams = Record<string, any> | undefined;

export type SPMetadataInfo = {
    type: string;
    id?: string;
    uri?: string;
};

export type SPMetadata = {
    __metadata: SPMetadataInfo;
};

export type SPDeferredInfo = {
    uri: string;
};

export type SPDeferred = {
    __deferred: SPDeferredInfo;
};

export type SPChangeToken = SPMetadata & {
    StringValue: string;
};

export type SPResourcePath = SPMetadata & {
    DecodedUrl: string;
};

export type SPResults<T> = {
    results: T[];
};

export type SPBase = SPMetadata & {
    Author: SPDeferred;
    ContentTypes: SPDeferred;
    Created: string;
    CurrentChangeToken: SPChangeToken;
    Description: string;
    DescriptionResource: SPDeferred;
    EventReceivers: SPDeferred;
    Fields: SPDeferred;
    FirstUniqueAncestorSecurableObject: SPDeferred;
    Id: string;
    LastItemModifiedDate: string;
    LastItemUserModifiedDate: string;
    NoCrawl: boolean;
    ParentWeb: SPDeferred;
    RoleAssignments: SPDeferred;
    RootFolder: SPDeferred;
    Title: string;
    TitleResource: SPDeferred;
    UserCustomActions: SPDeferred;
    WorkflowAssociations: SPDeferred;
};

export type SPList = SPBase & {
    AllowContentTypes: boolean;
    BaseTemplate: number;
    BaseType: number;
    ContentTypesEnabled: boolean;
    CrawlNonDefaultViews: boolean;
    CreatablesInfo: SPDeferred;
    DefaultContentApprovalWorkflowId: string;
    DefaultItemOpenUseListSetting: boolean;
    DefaultSensitivityLabelForLibrary: string;
    DefaultView: SPDeferred;
    Direction: string;
    DisableCommenting: boolean;
    DisableGridEditing: boolean;
    DocumentTemplateUrl: string | null;
    DraftVersionVisibility: number;
    EnableAttachments: boolean;
    EnableFolderCreation: boolean;
    EnableMinorVersions: boolean;
    EnableModeration: boolean;
    EnableRequestSignOff: boolean;
    EnableVersioning: boolean;
    EntityTypeName: string;
    ExemptFromBlockDownloadOfNonViewableFiles: boolean;
    FileSavePostProcessingEnabled: boolean;
    ForceCheckout: boolean;
    Forms: SPDeferred;
    HasExternalDataSource: boolean;
    Hidden: boolean;
    ImagePath: SPResourcePath;
    ImageUrl: string;
    InformationRightsManagementSettings: SPDeferred;
    IrmEnabled: boolean;
    IrmExpire: boolean;
    IrmReject: boolean;
    IsApplicationList: boolean;
    IsCatalog: boolean;
    IsPrivate: boolean;
    ItemCount: number;
    Items: SPDeferred;
    LastItemDeletedDate: string;
    ListExperienceOptions: number;
    ListItemEntityTypeFullName: string;
    MajorVersionLimit: number;
    MajorWithMinorVersionsLimit: number;
    MultipleDataList: boolean;
    ParentWebPath: SPResourcePath;
    ParentWebUrl: string;
    ParserDisabled: boolean;
    ServerTemplateCanCreateFolders: boolean;
    Subscriptions: SPDeferred;
    TemplateFeatureId: string;
    Views: SPDeferred;
};

export type SPWeb = SPBase & {
    AccessRequestsList: SPDeferred;
    Activities: SPDeferred;
    ActivityLogger: SPDeferred;
    Alerts: SPDeferred;
    AllowRssFeeds: boolean;
    AllProperties: SPDeferred;
    AlternateCssUrl: string;
    AppInstanceId: string;
    AppTiles: SPDeferred;
    AssociatedMemberGroup: SPDeferred;
    AssociatedOwnerGroup: SPDeferred;
    AssociatedVisitorGroup: SPDeferred;
    AvailableContentTypes: SPDeferred;
    AvailableFields: SPDeferred;
    CanModernizeHomepage: SPDeferred;
    ClassicWelcomePage: string | null;
    ClientWebParts: SPDeferred;
    Configuration: number;
    CurrentUser: SPDeferred;
    CustomMasterUrl: string;
    DataLeakagePreventionStatusInfo: SPDeferred;
    DesignPackageId: string;
    DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: boolean;
    EnableMinimalDownload: boolean;
    Features: SPDeferred;
    Folders: SPDeferred;
    FooterEmphasis: number;
    FooterEnabled: boolean;
    FooterLayout: number;
    HeaderEmphasis: number;
    HeaderLayout: number;
    HideTitleInHeader: boolean;
    HorizontalQuickLaunch: boolean;
    HostedApps: SPDeferred;
    IsEduClass: boolean;
    IsHomepageModernized: boolean;
    IsMultilingual: boolean;
    IsRevertHomepageLinkHidden: boolean;
    Language: number;
    Lists: SPDeferred | SPResults<SPList>;
    ListTemplates: SPDeferred;
    LogoAlignment: number;
    MasterUrl: string;
    MegaMenuEnabled: boolean;
    MultilingualSettings: SPDeferred;
    NavAudienceTargetingEnabled: boolean;
    Navigation: SPDeferred;
    ObjectCacheEnabled: boolean;
    OneDriveSharedItems: SPDeferred;
    OverwriteTranslationsOnChange: boolean;
    PushNotificationSubscribers: SPDeferred;
    QuickLaunchEnabled: boolean;
    RecycleBin: SPDeferred;
    RecycleBinEnabled: boolean;
    RegionalSettings: SPDeferred;
    ResourcePath: SPResourcePath;
    RoleDefinitions: SPDeferred;
    SearchScope: number;
    ServerRelativeUrl: string;
    SiteCollectionAppCatalog: SPDeferred;
    SiteGroups: SPDeferred;
    SiteLogoUrl: string | null;
    SiteUserInfoList: SPDeferred;
    SiteUsers: SPDeferred;
    SyndicationEnabled: boolean;
    TenantAdminMembersCanShare: number;
    TenantAppCatalog: SPDeferred;
    ThemeInfo: SPDeferred;
    TreeViewEnabled: boolean;
    UIVersion: number;
    UIVersionConfigurationEnabled: boolean;
    Url: string;
    WebInfos: SPDeferred;
    Webs: SPDeferred | SPResults<SPWeb>;
    WebTemplate: string;
    WelcomePage: string;
    WorkflowTemplates: SPDeferred;
};
