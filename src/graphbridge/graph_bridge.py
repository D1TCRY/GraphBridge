# pip install azure-identity requests
import json
import warnings
from copy import copy
from typing import Any
from urllib.parse import quote

from azure.identity import ClientSecretCredential

from .client import GraphBridgeClient
from .exceptions import GraphAuthenticationError, GraphRequestError
from .models import ListInfo, ListItem, SiteInfo


def _legacy_error_details(error: GraphAuthenticationError | GraphRequestError) -> str:
    """Format a modern Graph error for a legacy result.

    Args:
        error: Authentication or request error to format.
    """
    status_code = error.status_code
    response_text = error.response_text or str(error)
    return f"{status_code} {response_text}"


def deduplicate_dicts(dict_list: list[dict]) -> list[dict]:
    """Remove duplicate dictionaries while preserving order.

    Dictionaries are compared through a stable JSON representation, so the first
    occurrence of each equivalent mapping is retained.

    Args:
        dict_list: Dictionaries to deduplicate.
    """
    unique_items = []
    seen_keys = []

    for item in dict_list:
        key = json.dumps(item, sort_keys=True)
        if key not in seen_keys:
            seen_keys.append(key)
            unique_items.append(item)
    
    return unique_items




class GbAuth(object):
    """Provide the legacy app-only authentication interface.

    This compatibility class retains lazy cached credentials and tokens expected
    by existing users while its HTTP operations can share the composed client.

    Args:
        tenant_id: Microsoft Entra tenant identifier.
        client_id: Application client identifier.
        client_secret: Application client secret.
    """

    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        """Initialize legacy authentication settings.

        Values are validated through their property setters. Credential creation
        and token acquisition remain lazy until an authenticated operation occurs.

        Args:
            tenant_id: Microsoft Entra tenant identifier.
            client_id: Application client identifier.
            client_secret: Application client secret.
        """
        self.__auth_name = "GbAuth"
        self.__credential: ClientSecretCredential
        self.__token: str
        self.__graph_client: GraphBridgeClient
        
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        
        # Autenticazione app-only con ClientSecretCredential
        # self.credential
        # self.token
        # self.headers
    
    def __str__(self) -> str:
        """Return a readable representation with the secret redacted."""
        return f"< {self.__auth_name} | Tenant ID: {self.tenant_id}, Client ID: {self.client_id}, Client Secret: <redacted> >"
    
    def __repr__(self) -> str:
        """Return a constructor-style representation with the secret redacted."""
        return (
            f"{self.__auth_name}(tenant_id={self.tenant_id!r}, client_id={self.client_id!r}, "
            "client_secret='<redacted>')"
        )
    
    @property
    def tenant_id(self) -> str:
        """Return the Microsoft Entra tenant identifier.

        The value is used when the legacy credential is created lazily.
        """
        return self.__tenant_id
    @tenant_id.setter
    def tenant_id(self, value: str) -> None:
        """Set the tenant identifier and clear authentication caches.

        Invalidating cached state ensures future tokens use the new tenant.

        Args:
            value: Non-empty tenant identifier.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__auth_name} | Tenant ID cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__auth_name} | Tenant ID must be string>")
        self.__tenant_id: str = value
        self._invalidate_auth_cache()
    
    @property
    def client_id(self) -> str:
        """Return the application client identifier.

        The value identifies the Entra application used for app-only access.
        """
        return self.__client_id
    @client_id.setter
    def client_id(self, value) -> None:
        """Set the client identifier and clear authentication caches.

        Future credential and token access will use the replacement value.

        Args:
            value: Non-empty client identifier.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__auth_name} | Client ID cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__auth_name} | Client ID must be string>")
        self.__client_id: str = value
        self._invalidate_auth_cache()
    
    @property
    def client_secret(self) -> str:
        """Return the application client secret.

        String representations deliberately redact this sensitive value.
        """
        return self.__client_secret
    @client_secret.setter
    def client_secret(self, value: str) -> None:
        """Set the client secret and clear authentication caches.

        Any cached credential, token, or composed client is discarded immediately.

        Args:
            value: Non-empty application client secret.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__auth_name} | Client Secret cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__auth_name} | Client Secret must be string>")
        self.__client_secret: str = value
        self._invalidate_auth_cache()
    
    @property
    def credential(self) -> ClientSecretCredential:
        """Return the cached Azure AD client-secret credential.

        The credential is created lazily from the configured tenant, client, and
        secret values and is invalidated when any of them changes.
        """
        
        if f"_{self.__auth_name}__credential" in self.__dict__:
            return self.__credential
        else:
            self.__credential = ClientSecretCredential(self.tenant_id, self.client_id, self.client_secret)
            return self.__credential
    
    @property
    def token(self) -> str:
        """Return the cached access token used by the legacy API.

        Unlike the composed transport, this property intentionally preserves the
        original legacy behavior of caching the token string after acquisition.
        """
        
        if f"_{self.__auth_name}__token" in self.__dict__:
            return self.__token
        else:
            try:
                self.__token = self.credential.get_token('https://graph.microsoft.com/.default').token
                return self.__token
            except Exception as e:
                try:
                    safe_cause = type(e)("Credential failure details redacted")
                except Exception:
                    safe_cause = RuntimeError("Credential failure details redacted")
                raise RuntimeError(f"<ERROR {self.__auth_name} | Failed to get token>") from safe_cause
    
    @property
    def headers(self) -> dict:
        """Return the legacy Microsoft Graph authorization headers.

        Accessing this property may trigger lazy credential and token acquisition.
        """
        # Preparazione dell'header di autorizzazione per le richieste Graph
        return {'Authorization': f'Bearer {self.token}'}

    def _invalidate_auth_cache(self) -> None:
        """Clear cached credential, token, and composed client values."""
        self.__dict__.pop("_GbAuth__credential", None)
        self.__dict__.pop("_GbAuth__token", None)
        self.__dict__.pop("_GbAuth__graph_client", None)

    def _get_graph_client(self) -> GraphBridgeClient:
        """Return the cached composed GraphBridge client."""
        if "_GbAuth__graph_client" not in self.__dict__:
            self.__graph_client = GraphBridgeClient(credential=self.credential, max_retries=0)
        return self.__graph_client

    def _adopt_graph_client(self, client: GraphBridgeClient) -> None:
        """Reuse an existing composed GraphBridge client.

        Args:
            client: Client to share with this legacy object.
        """
        self.__graph_client = client
    
    
    

class GbSite(GbAuth):
    """Provide the legacy SharePoint site interface.

    A site can be constructed from authentication values or an existing
    :class:`GbAuth`, in which case the underlying composed client is reused.

    Args:
        *args: Positional authentication arguments used without ``gb_auth``.
        hostname: SharePoint tenant hostname.
        site_path: Server-relative site path.
        gb_auth: Optional existing legacy authentication object.
        **kwargs: Keyword authentication arguments used without ``gb_auth``.
    """
    
    def __init__(self, *args, hostname: str, site_path: str, gb_auth: GbAuth | None = None, **kwargs) -> None:
        """Initialize the legacy site interface.

        Site metadata remains lazy and is fetched only when an ID or data property
        is accessed.

        Args:
            *args: Positional authentication arguments used without ``gb_auth``.
            hostname: SharePoint tenant hostname.
            site_path: Server-relative site path.
            gb_auth: Optional existing legacy authentication object.
            **kwargs: Keyword authentication arguments used without ``gb_auth``.

        Raises:
            TypeError: If ``gb_auth`` is not a ``GbAuth`` instance.
        """
        self.__site_name = "GbSite"
        self.__site_data: dict
        
        self.hostname = hostname
        self.site_path = site_path
        # site_url
        # site_data
        
        if gb_auth is None:
            super().__init__(*args, **kwargs)
        else:
            if not isinstance(gb_auth, GbAuth):
                raise TypeError(f"<ERROR {self.__site_name} | sp_auth must be an instance of 'SPAuth'>")
            self.__gb_auth = gb_auth
            super().__init__(tenant_id=gb_auth.tenant_id, client_id=gb_auth.client_id, client_secret=gb_auth.client_secret)
            self._adopt_graph_client(gb_auth._get_graph_client())
    
    def __str__(self) -> str:
        """Return a readable representation of the site."""
        return f"< {self.__site_name} | Hostname: {self.hostname}, Site Path: {self.site_path}, Site ID: {self.site_id} >"
    
    @property
    def hostname(self) -> str:
        """Return the SharePoint site hostname.

        The hostname forms the tenant portion of the legacy site-resolution URL.
        """
        return self.__hostname
    @hostname.setter
    def hostname(self, value: str) -> None:
        """Set the SharePoint hostname.

        This preserves the characterized legacy behavior and does not invalidate
        site metadata that may already be cached.

        Args:
            value: Non-empty SharePoint hostname.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__site_name} | Hostname cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__site_name} | Hostname must be string>")
        self.__hostname: str = value
        
    @property
    def site_path(self) -> str:
        """Return the server-relative SharePoint site path.

        The path is combined with the hostname when resolving site metadata.
        """
        
        return self.__site_path
    @site_path.setter
    def site_path(self, value: str) -> None:
        """Set the server-relative site path.

        This preserves the characterized legacy behavior and does not invalidate
        metadata already stored on the object.

        Args:
            value: Non-empty site path.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__site_name} | Site path cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__site_name} | Site path must be string>")
        self.__site_path: str = value
    
    @property
    def site_url(self) -> str:
        """Construct the SharePoint site-resolution URL.

        The legacy property returns an absolute stable Graph v1.0 URL.
        """
        
        return f"https://graph.microsoft.com/v1.0/sites/{self.hostname}:{self.site_path}"
    
    @property
    def site_data(self) -> dict:
        """Fetch and cache the SharePoint site metadata.

        Modern typed transport errors are wrapped in the ``RuntimeError`` shape
        expected by the characterized legacy contract.
        """
        
        if "_GbSite__site_data" in self.__dict__:
            return self.__site_data
        else:
            try:
                self.__site_data = self._get_graph_client().transport.get(self.site_url)
            except (GraphAuthenticationError, GraphRequestError) as error:
                raise RuntimeError(
                    f"<ERROR {self.__site_name} | Failed to set site data: {_legacy_error_details(error)}>"
                ) from error
            return self.__site_data
    
    @property
    def site_id(self) -> str:
        """Return the SharePoint site identifier.

        Accessing the property triggers lazy metadata retrieval and returns the
        documented warning string if Graph omitted the ID.
        """
        return self.site_data.get("id", "<WARNING SPM | Site ID not found>")
    
    
    
    
class GbList(GbSite):
    """Provide the legacy SharePoint list interface.

    The class retains dictionary result shapes, cached metadata, local filtering,
    heuristic field encoding, and legacy batch helpers during migration.

    Args:
        *args: Positional site or authentication arguments used without ``gb_site``.
        list_name: SharePoint list name or title.
        gb_site: Optional existing legacy site object.
        **kwargs: Keyword site or authentication arguments used without ``gb_site``.
    """
    
    def __init__(self, *args, list_name: str, gb_site: GbSite | None = None, **kwargs) -> None:
        """Initialize the legacy list interface.

        An existing :class:`GbSite` may be injected to reuse its location,
        credentials, and composed client; list metadata itself remains lazy.

        Args:
            *args: Positional site or authentication arguments used without ``gb_site``.
            list_name: SharePoint list name or title.
            gb_site: Optional existing legacy site object.
            **kwargs: Keyword site or authentication arguments used without ``gb_site``.

        Raises:
            TypeError: If ``gb_site`` is not a ``GbSite`` instance.
        """
        self.__list_obj_name = "GbList"
        self.__list_data: dict
        
        self.list_name = list_name
        
        if gb_site is None:
            super().__init__(*args, **kwargs)
        else:
            if not isinstance(gb_site, GbSite):
                raise TypeError(f"<ERROR {self.__list_obj_name} | gb_site must be an instance of 'GbSite'>")
            self.__gb_site = gb_site
            super().__init__(tenant_id=gb_site.tenant_id, client_id=gb_site.client_id, client_secret=gb_site.client_secret, 
                            hostname=gb_site.hostname, site_path=gb_site.site_path)
            self._adopt_graph_client(gb_site._get_graph_client())
    
    def __str__(self) -> str:
        """Return a readable representation of the list."""
        return f"< {self.__list_obj_name} | list_name: {self.list_name}, list_id: {self.list_id} >"
    
    def __repr__(self) -> str:
        """Return a safe constructor-style representation of the list."""
        return (
            f"{self.__list_obj_name}(list_name={self.list_name!r}, hostname={self.hostname!r}, "
            f"site_path={self.site_path!r}, tenant_id={self.tenant_id!r}, client_id={self.client_id!r}, "
            "client_secret='<redacted>')"
        )
    
    @property
    def encode_map(self) -> dict:
        """Return the legacy character-to-field-code mapping.

        The table approximates SharePoint internal-name encoding and is retained
        only for backward compatibility.
        """
        self.__encode_map = {
            # Spazio e punteggiatura comune
            ' ': '_x0020_',
            '/': '_x002f_',
            '.': '_x002e_',
            ',': '_x002c_',
            ':': '_x003a_',
            ';': '_x003b_',
            '!': '_x0021_',
            '?': '_x003f_',
            '@': '_x0040_',
            '#': '_x0023_',
            '$': '_x0024_',
            '%': '_x0025_',
            '&': '_x0026_',
            '(': '_x0028_',
            ')': '_x0029_',
            '+': '_x002b_',
            '-': '_x002d_',
            '=': '_x003d_',
            "'": '_x0027_',
            '"': '_x0022_',
            '\\': '_x005c_',
            '*': '_x002a_',

            # Numeri da 0 a 9
            '0': '_x0030_',
            '1': '_x0031_',
            '2': '_x0032_',
            '3': '_x0033_',
            '4': '_x0034_',
            '5': '_x0035_',
            '6': '_x0036_',
            '7': '_x0037_',
            '8': '_x0038_',
            '9': '_x0039_',

            # Lettere accentate (facoltativo, se usi nomi internazionali)
            'à': '_x00e0_',
            'è': '_x00e8_',
            'é': '_x00e9_',
            'ì': '_x00ec_',
            'ò': '_x00f2_',
            'ù': '_x00f9_',
            'ç': '_x00e7_',

            # Simboli matematici e vari
            '<': '_x003c_',
            '>': '_x003e_',
            '[': '_x005b_',
            ']': '_x005d_',
            '{': '_x007b_',
            '}': '_x007d_',
            '^': '_x005e_',
            '`': '_x0060_',
            '|': '_x007c_',
            '~': '_x007e_',
        }
        return self.__encode_map
    
    @property
    def decode_map(self) -> dict:
        """Return the inverse legacy field-name mapping.

        It is generated from ``encode_map`` so encoding and decoding remain paired.
        """
        
        self.__decode_map = {v: k for k, v in self.encode_map.items()}
        return self.__decode_map
    
    def decode_row(self, row: dict) -> dict:
        """Decode legacy SharePoint field names in a row.

        Every encoded substring known to ``decode_map`` is replaced in each key;
        values are preserved unchanged.

        Args:
            row: Field mapping whose keys should be decoded.
        """
        
        decoded_row = {}
        for k, v in row.items():
            decoded_k = copy(k)
            for char, decoded_char in self.decode_map.items():
                if char in k:
                    decoded_k = decoded_k.replace(char, decoded_char)
            decoded_row[decoded_k] = v
        return decoded_row
    
    def encode_row(self, row: dict) -> dict:
        """Encode field names with the legacy character map.

        This heuristic is retained for compatibility only and does not verify the
        actual internal names stored in the list schema.

        Args:
            row: Field mapping whose keys should be encoded.
        """
        
        encoded_row = {}
        for k, v in row.items():
            encoded_k = copy(k)
            for char, encoded_char in self.encode_map.items():
                if char in k:
                    encoded_k = encoded_k.replace(char, encoded_char)
            encoded_row[encoded_k] = v
        return encoded_row
    
    @property
    def list_name(self) -> str:
        """Return the configured SharePoint list name.

        The value is used as a quoted title segment in the legacy list URL.
        """
        return self.__list_name
    @list_name.setter
    def list_name(self, value: str) -> None:
        """Set the SharePoint list name.

        Changing it preserves the characterized legacy behavior and does not clear
        list metadata that may already be cached.

        Args:
            value: Non-empty list name or title.

        Raises:
            TypeError: If the value is not a string.
            ValueError: If the value is empty.
        """
        if not value:
            raise ValueError(f"<ERROR {self.__list_obj_name} | Name cannot be empty>")
        elif not isinstance(value, str):
            raise TypeError(f"<ERROR {self.__list_obj_name} | Name must be string>")
        self.__list_name: str = value
    
    @property
    def list_url(self) -> str:
        """Construct the stable Graph URL for the SharePoint list.

        The configured list name is quoted and combined with the lazily resolved
        site identifier.
        """
        
        return f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{quote(self.list_name)}"
    
    @property
    def list_data(self) -> dict:
        """Fetch and cache the SharePoint list metadata.

        The request uses the configured title path and wraps typed transport errors
        in the legacy ``RuntimeError`` contract.
        """
        
        if f"_{self.__list_obj_name}__list_data" in self.__dict__:
            return self.__list_data
        else:
            try:
                self.__list_data = self._get_graph_client().transport.get(self.list_url)
            except (GraphAuthenticationError, GraphRequestError) as error:
                raise RuntimeError(
                    f"<ERROR {self.__list_obj_name} | Failed to set element data: {_legacy_error_details(error)}>"
                ) from error
            return self.__list_data
        
    @property
    def list_id(self) -> str:
        """Return the SharePoint list identifier.

        Access triggers lazy metadata retrieval and falls back to the documented
        warning string when the response omits an ID.
        """
        
        return self.list_data.get("id", f"<WARNING {self.__list_obj_name} | Element ID not found>")
    
    @property
    def list_items_all(self, top: int = 200) -> list[dict]:
        """Return all list items across every page.

        The method eagerly follows every continuation link and preserves the raw
        legacy dictionary shape rather than returning typed item models.

        Args:
            top: Requested page size for the initial Graph call.

        Raises:
            RuntimeError: If authentication or the Graph request fails.
        """
        url = f"{self.list_url}/items?expand=fields&$top={top}"
        items: list[dict] = []
        while url:
            try:
                data = self._get_graph_client().transport.get(url)
            except (GraphAuthenticationError, GraphRequestError) as error:
                raise RuntimeError(
                    f"<ERROR {self.__list_obj_name} | Failed to fetch list items: {_legacy_error_details(error)}>"
                ) from error
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # se presente, contiene l’URL della pagina successiva
        return items
    
    @property
    def list_items(self) -> list:
        """Fetch the first page of items in the SharePoint list.

        Despite its historical name, this property performs one collection call;
        ``list_items_all`` is the paginated variant.
        """
        
        items_url = f"{self.list_url}/items?expand=fields"
        try:
            data = self._get_graph_client().transport.get(items_url)
        except (GraphAuthenticationError, GraphRequestError) as error:
            raise RuntimeError(
                f"<ERROR {self.__list_obj_name} | Failed to fetch list items: {_legacy_error_details(error)}>"
            ) from error
        self.__list_items = data.get("value", [])
        return self.__list_items
    
    @property
    def list_rows(self) -> list[dict]:
        """Return field dictionaries for all paginated list items.

        Item metadata is discarded and only each item's ``fields`` mapping is
        retained in the legacy result.
        """
        
        self.__list_rows = [item["fields"] for item in self.list_items_all]
        return self.__list_rows
    
    @property
    def list_ids(self) -> list[str]:
        """Return identifiers from all paginated list items.

        Entries without an ``id`` property are omitted to preserve the legacy
        collection shape.
        """
        
        return [item.get("id", "None") for item in self.list_items_all if "id" in item]
    
    @property
    def list_fields(self) -> list[str]:
        """Return field names observed on the first list row.

        An empty list produces an empty result, and heterogeneous later rows are
        not inspected for additional fields.
        """
        
        rows = self.list_rows
        return list(rows[0].keys()) if rows else []
    
    
    def get_items_by_features(self, features: list[dict]) -> list[dict[str, object]]:
        """
        Return SharePoint list items that match at least ONE of the provided feature
        sets (logical OR across the dicts in `features`). Within each single feature
        dict, all key/value pairs must match (logical AND). One level of nesting is
        supported (e.g., {"Field": {"SubField": value}}).

        Args:
            features (list[dict]): A list of criteria dictionaries. Each dict may
                contain:
                - flat pairs {field: value} for direct comparisons; and/or
                - nested pairs {field: {sub_field: value}} for nested comparisons.

        Returns:
            list[dict]: A de-duplicated list of items that satisfy at least one dict
            in `features`.

        Example:
            features = [
                {"Status": "Active"},
                {"Category": {"Name": "Premium"}}
            ]
            items = obj.get_items_by_features(features)
        """

        
        items = []
        for feature in features:
            for item in self.list_items_all:
                has_features = []
                for k, v in feature.items():
                    if not isinstance(v, dict):
                        has_features.append(item.get(k) == v)
                    else:
                        nested_has_features = []
                        for nested_k, nested_v in v.items():
                            nested_has_features.append(item.get(k, {}).get(nested_k) == nested_v)
                        has_features.append(all(nested_has_features))
                
                if all(has_features):
                    items.append(item)
            
            # if all(item.get(k) == v if not isinstance(v, dict) else all(item.get(k, {}).get(nested_k) == nested_v for nested_k, nested_v in v.items()) for k, v in feature.items()):
            #     items.append(item)
        return deduplicate_dicts(items)
    
    
    def update(self, ids: str | int | list[str] | tuple[str] | set[str], rows: dict | list[dict] | tuple[dict]) -> dict[str, list[dict[str, object]]]:
        """
        Update one or more SharePoint list items by their IDs.

        IDs and rows can be single values or sequences; the method normalizes them to
        lists and requires a 1:1 correspondence (same length).

        Args:
            ids (str | int | list[str] | tuple[str] | set[str]): A single ID or a
                collection of IDs. Single values are accepted and will be converted.
            rows (dict | list[dict] | tuple[dict]): Update payload (field/value pairs)
                for each ID. A single dict is accepted and will be converted to a list.

        Returns:
            dict: A result object with:
                - "successes": list of {"id", "success", "updated_row"}
                - "failures":  list of {"id", "success", "error"}

        Raises:
            TypeError: If `rows` or `ids` cannot be converted to a list as required.
            ValueError: If the number of `ids` does not match the number of `rows`.

        Note:
            Although the type hint says `-> bool`, this function actually returns a
            result dictionary with "successes"/"failures".
        """
        
        if isinstance(rows, dict):
            normalized_rows = [rows]
        elif not isinstance(rows, list):
            try:
                normalized_rows = list(rows)
            except TypeError:
                raise TypeError(f"<ERROR {self.__list_obj_name} | rows must be a list, tuple or dict>")
        else:
            normalized_rows = rows
        
        normalized_ids: list[str | int]
        if isinstance(ids, str | int):
            normalized_ids = [ids]
        elif not isinstance(ids, list):
            try:
                normalized_ids = [item for item in ids]
            except TypeError:
                raise TypeError(f"<ERROR {self.__list_obj_name} | IDs must be a list, tuple, set or string>")
        else:
            normalized_ids = [item for item in ids]
        
        if len(normalized_ids) != len(normalized_rows):
            raise ValueError(f"<ERROR {self.__list_obj_name} | Number of IDs must match number of rows to update>")
        
        update_successes: list = []
        update_failures: list = []
        for i, id in enumerate(normalized_ids):
            row = normalized_rows[i]
            
            update_url: str = f"{self.list_url}/items/{id}/fields"
            try:
                updated_row: dict = self._get_graph_client().transport.patch(update_url, json=row)
                update_successes.append({"id": updated_row.get("id", None), "success": True, "updated_row": updated_row})
            except (GraphAuthenticationError, GraphRequestError) as error:
                update_failures.append({"id": id, "success": False, "error": f"Error updating: {_legacy_error_details(error)}"})
        
        return {"successes": update_successes, "failures": update_failures}
        
        
    def create(self, rows: dict | list[dict] | tuple[dict] | set[dict]) -> dict[str, list[dict[str, object]]]:
        """
        Create one or more new items in the SharePoint list.

        Each row is wrapped as {"fields": row} before being sent to the creation
        endpoint.

        Args:
            rows (dict | list[dict] | tuple[dict] | set[dict]): Field data for new
                items. You may pass a single dict or a collection of dicts; non-list
                inputs are converted.

        Returns:
            dict: A result object with:
                - "successes": list of {"id", "success", "item"}
                - "failures":  list of {"success", "error"}

        Raises:
            TypeError: If any row is not a dict, or if `rows` cannot be converted to a
                collection of dicts.

        Note:
            The return annotation is `dict[str, dict]`, but the actual return value is
            a dictionary containing lists of outcome records.
        """

        if isinstance(rows, dict):
            rows = [{"fields": rows}]
        elif not isinstance(rows, list):
            try:
                rows = list(rows)
                rows = [{"fields": row} for row in rows]
            except TypeError:
                raise TypeError(f"<ERROR {self.__list_obj_name} | rows must be a list, tuple, set or dict>")
        else:
            rows = [{"fields": row} for row in rows]

        for new_row in rows:
            if not isinstance(new_row, dict):
                raise TypeError(f"<ERROR {self.__list_obj_name} | Each row must be a dictionary>")
        
        create_successes: list = []
        create_failures: list = []
        for new_row in rows:
            create_url: str = self.list_url + "/items"
            try:
                item_created: dict = self._get_graph_client().transport.post(create_url, json=new_row)
                create_successes.append({"id": item_created.get("id", None), "success": True, "item": item_created})
            except (GraphAuthenticationError, GraphRequestError) as error:
                create_failures.append({"success": False, "error": f"Error while creating a new item: {_legacy_error_details(error)}"})
        
        return {"successes": create_successes, "failures": create_failures}


    def delete(self, ids: str | list[str] | tuple[str] | set[str]) -> dict[str, list[dict[str, object]]]:
        """
        Delete one or more SharePoint list items by ID.

        Every item is deleted independently, allowing the legacy result to report
        partial success without raising for ordinary Graph request failures.

        Args:
            ids (str | list[str] | tuple[str] | set[str]): A single ID or a collection
                of IDs. Single values are converted to a list.

        Returns:
            dict: A result object with:
                - "successes": list of {"id", "completed", "message"}
                - "failures":  list of {"id", "completed", "error"}

        Raises:
            TypeError: If `ids` cannot be converted to a list.
            ValueError: If `ids` is empty.

        Note:
            The delete endpoint returns HTTP 204 on success.
        """

        if isinstance(ids, str):
            ids = [ids]
        elif not isinstance(ids, list):
            try:
                ids = list(ids)
            except TypeError:
                raise TypeError(f"<ERROR {self.__list_obj_name} | IDs must be a list, tuple, set or string>")
        
        if not ids:
            raise ValueError(f"<ERROR {self.__list_obj_name} | IDs cannot be empty>")
        
        del_successes: list = []
        del_failures: list = []
        for id in ids:
            del_url: str = f"{self.list_url}/items/{id}"
            try:
                self._get_graph_client().transport.delete(del_url)
                del_successes.append({"id": id, "completed": True, "message": "Item deleted successfully."})
            except (GraphAuthenticationError, GraphRequestError) as error:
                del_failures.append({"id": id, "completed": False, "error": f"Error while deleting: {_legacy_error_details(error)}"})
        
        return {"successes": del_successes, "failures": del_failures}
    
    
    def upload(self, ids: str | int | list[str | int] | tuple[str | int] | set[str | int], rows: dict | list[dict] | tuple[dict], force: bool = False, delete: bool = False) -> dict[str, Any]:
        """Apply the deprecated legacy upload workflow.

        The adapter builds a modern synchronization plan but maps every outcome
        back to the historical result sections. ``force`` performs PATCH updates,
        never delete-and-recreate, and pruning runs after successful creates.

        Args:
            ids: Source item identifiers paired with rows.
            rows: Source field mappings.
            force: Whether updates appear in the legacy ``replaced`` section.
            delete: Whether remote-only items should be pruned.

        Raises:
            TypeError: If an input has an unsupported type.
            ValueError: If identifier and row counts differ.
        """

        warnings.warn(
            "GbList.upload() is deprecated; use tasks.sync.plan() and "
            "tasks.sync.apply() to review changes before applying them",
            DeprecationWarning,
            stacklevel=2,
        )
        if isinstance(ids, str | int):
            normalized_ids = [str(ids)]
        else:
            try:
                normalized_ids = [str(item_id) for item_id in ids]
            except TypeError:
                raise TypeError(
                    f"<ERROR {self.__list_obj_name} | IDs must be a list, tuple, set, string or integer>"
                ) from None
        if isinstance(rows, dict):
            normalized_rows = [rows]
        elif not isinstance(rows, list):
            try:
                normalized_rows = list(rows)
            except TypeError:
                raise TypeError(
                    f"<ERROR {self.__list_obj_name} | rows must be a list, tuple or dict>"
                ) from None
        else:
            normalized_rows = rows
        if not all(isinstance(row, dict) for row in normalized_rows):
            raise TypeError(
                f"<ERROR {self.__list_obj_name} | Each row must be a dictionary>"
            )
        if len(normalized_ids) != len(normalized_rows):
            raise ValueError(
                f"<ERROR {self.__list_obj_name} | Number of IDs must match number of rows to upload>"
            )
        if not isinstance(force, bool):
            raise TypeError(f"<ERROR {self.__list_obj_name} | force must be a boolean>")
        if not isinstance(delete, bool):
            raise TypeError(f"<ERROR {self.__list_obj_name} | delete must be a boolean>")

        raw_remote_items = self.list_items_all
        remote_items = [
            ListItem.from_payload(item)
            for item in raw_remote_items
            if isinstance(item, dict)
        ]
        client = self._get_graph_client()
        site = client.sites.bind(SiteInfo(id=self.site_id))
        tasks = site.lists.bind(
            ListInfo(id=self.list_name, display_name=self.list_name)
        )
        plan = tasks.sync._plan_by_item_ids(
            rows=normalized_rows,
            item_ids=normalized_ids,
            remote_items=remote_items,
            prune=delete,
        )
        sync_result = tasks.sync.apply(plan, use_batch=False)

        delete_results: dict[str, Any] = {
            "successes": [] if delete else None,
            "failures": [] if delete else None,
        }
        force_results: dict[str, dict[str, list[dict[str, Any]]]] = {
            "replaced": {"successes": [], "failures": []},
            "updated": {"successes": [], "failures": []},
            "created": {"successes": [], "failures": []},
        }
        for outcome in sync_result.results:
            operation = outcome.operation
            if operation.operation == "unchanged" and operation.source_index is None:
                continue
            row = (
                normalized_rows[operation.source_index]
                if operation.source_index is not None
                else dict(operation.fields)
            )
            if operation.operation == "delete":
                assert delete_results["successes"] is not None
                assert delete_results["failures"] is not None
                if outcome.succeeded:
                    delete_results["successes"].append(
                        {
                            "id": operation.item_id,
                            "completed": True,
                            "message": "Item deleted successfully.",
                        }
                    )
                else:
                    delete_results["failures"].append(
                        {
                            "id": operation.item_id,
                            "completed": False,
                            "error": self._legacy_sync_error(outcome.error),
                        }
                    )
            elif operation.operation == "create":
                target = force_results["created"]
                if outcome.succeeded:
                    value = outcome.value
                    target["successes"].append(
                        {
                            "id": operation.key,
                            "row": row,
                            "new_id": value.id if isinstance(value, ListItem) else None,
                        }
                    )
                else:
                    target["failures"].append(
                        {
                            "id": operation.key,
                            "row": row,
                            "error": self._legacy_sync_error(outcome.error),
                        }
                    )
            else:
                target = force_results["replaced" if force else "updated"]
                if outcome.succeeded:
                    entry: dict[str, Any] = {"id": operation.key, "row": row}
                    if force:
                        entry["new_id"] = operation.item_id
                    target["successes"].append(entry)
                else:
                    target["failures"].append(
                        {
                            "id": operation.key,
                            "row": row,
                            "error": self._legacy_sync_error(outcome.error),
                        }
                    )
        return {"delete_results": delete_results, "force_results": force_results}

    @staticmethod
    def _legacy_sync_error(error: object) -> str:
        """Format a synchronization error for legacy output.

        Args:
            error: Structured error or arbitrary failure value.
        """
        if error is None:
            return "Synchronization operation failed"
        status = getattr(error, "status_code", None)
        message = getattr(error, "message", str(error))
        return f"{status} {message}" if status is not None else str(message)


    def create_many(self, rows: list[dict], batch_size: int = 20) -> dict:
        """Create multiple items with the legacy batch interface.

        Rows are divided into caller-sized batches and returned as legacy success
        and failure dictionaries rather than typed batch outcomes.

        Args:
            rows: Field mappings for the items to create.
            batch_size: Maximum number of items per batch request.

        Raises:
            TypeError: If rows is not a list of dictionaries.
        """
        if not isinstance(rows, list) or not all(isinstance(r, dict) for r in rows):
            raise TypeError("rows deve essere una lista di dict")

        def chunked(seq, size):
            """Split a sequence into fixed-size chunks.

            Chunks preserve input order for legacy batch correlation.

            Args:
                seq: Sequence to split.
                size: Maximum values per chunk.
            """
            for i in range(0, len(seq), size):
                yield seq[i:i+size]

        results: dict[str, list[dict[str, Any]]] = {"successes": [], "failures": []}
        for chunk_idx, chunk in enumerate(chunked(rows, batch_size), start=1):
            batch_body: dict[str, Any] = {
                "requests": [
                    {
                        "id": f"{chunk_idx}-{i+1}",
                        "method": "POST",
                        "url": f"/sites/{self.site_id}/lists/{self.list_id}/items",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"fields": row}
                    }
                    for i, row in enumerate(chunk)
                ]
            }

            try:
                data = self._get_graph_client().transport.post(
                    "https://graph.microsoft.com/v1.0/$batch",
                    json=batch_body,
                )
            except (GraphAuthenticationError, GraphRequestError) as error:
                results["failures"].append({
                    "batch": chunk_idx,
                    "error": _legacy_error_details(error)
                })
                continue

            for r in data.get("responses", []):
                if 200 <= r.get("status", 0) < 300:
                    # Il body contiene l'item creato come da API create listItem
                    results["successes"].append({
                        "id": r.get("body", {}).get("id"),
                        "status": r["status"],
                        "item": r.get("body")
                    })
                else:
                    results["failures"].append({
                        "id": r.get("id"),
                        "status": r.get("status"),
                        "error": r.get("body")
                    })
        return results


    def delete_many(self, ids, batch_size: int = 20, if_match: str | None = None) -> dict:
        """Delete multiple items with the legacy batch interface.

        Each subrequest is correlated to its item ID and may carry the same
        optional ``If-Match`` value for optimistic concurrency.

        Args:
            ids: One item ID or an iterable of item IDs.
            batch_size: Maximum number of items per batch request.
            if_match: Optional eTag sent with each delete.

        Raises:
            TypeError: If IDs cannot be normalized to a list.
        """
        # normalizza gli id
        if isinstance(ids, (str, int)):
            ids = [str(ids)]
        elif not isinstance(ids, list):
            try:
                ids = [str(x) for x in ids]
            except TypeError:
                raise TypeError("ids deve essere stringa, intero o iterabile di valori")

        def chunked(seq, size):
            """Split a sequence into fixed-size chunks.

            Chunks preserve input order for legacy batch correlation.

            Args:
                seq: Sequence to split.
                size: Maximum values per chunk.
            """
            for i in range(0, len(seq), size):
                yield seq[i:i+size]

        results: dict[str, list[dict[str, Any]]] = {"successes": [], "failures": []}

        for batch_idx, chunk in enumerate(chunked(ids, batch_size), start=1):
            # costruisci le sottorichieste DELETE
            requests_map = {}
            reqs = []
            for i, item_id in enumerate(chunk, start=1):
                req_id = f"{batch_idx}-{i}"
                requests_map[req_id] = item_id
                req: dict[str, Any] = {
                    "id": req_id,
                    "method": "DELETE",
                    "url": f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}",
                }
                if if_match is not None:
                    req["headers"] = {"If-Match": if_match}
                reqs.append(req)

            body = {"requests": reqs}
            try:
                data = self._get_graph_client().transport.post(
                    "https://graph.microsoft.com/v1.0/$batch",
                    json=body,
                )
            except (GraphAuthenticationError, GraphRequestError) as error:
                results["failures"].append({
                    "batch": batch_idx,
                    "error": _legacy_error_details(error),
                })
                continue

            for r in data.get("responses", []):
                status = r.get("status")
                rid = r.get("id")
                item_id = requests_map.get(rid)
                if status == 204:
                    results["successes"].append({"id": item_id, "status": status})
                else:
                    results["failures"].append({
                        "id": item_id,
                        "status": status,
                        "error": r.get("body"),
                    })

        return results
