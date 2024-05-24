import logging
import re
from typing import List, Optional

import httpx
import jwt
from aiocache import Cache, cached
from fastapi import Depends, Header, status
from fastapi.exceptions import HTTPException
from pydantic import BaseModel

from ..config import settings

logger = logging.getLogger(__name__)

OPENID_CONNECT_DISCOVERY_DOCUMENT_URL = (
    "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"
)


class User(BaseModel):
    id: str
    name: str
    email: Optional[str] = None
    roles: Optional[List[str]] = []


@cached(ttl=60 * 60 * 24, cache=Cache.MEMORY)
async def get_jwks_client_and_algorithms():
    async with httpx.AsyncClient() as client:
        response = await client.get(OPENID_CONNECT_DISCOVERY_DOCUMENT_URL)
    data = response.json()
    jwks_uri = data["jwks_uri"]
    algorithms = data["id_token_signing_alg_values_supported"]
    jwks_client = jwt.PyJWKClient(jwks_uri)  # TODO: Makes a blocking request
    return jwks_client, algorithms


@cached(ttl=60 * 60, cache=Cache.MEMORY)
async def validate_token(token: str):
    """Function that reads and validates the Entra ID access/id token.
    Returns a user object."""
    if not settings.entraid_tenant_id and not settings.entraid_client_id:
        return User(id="n/a", name="Anonymous")
    logger.debug(f"Validating token: {token}")
    if token.lower().startswith("error"):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Auth error with Entra ID token: {token} See https://learn.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins#causes-and-handling-of-errors-from-getaccesstoken",
        )
    if token.lower().startswith("bearer"):
        parts = token.split()
        if len(parts) != 2:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Auth error: Invalid token, must be in format 'Bearer xxxx'",
            )
        token = parts[1]
    else:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Auth error: Invalid token, must be in format 'Bearer xxxx'",
        )
    jwks_client, algorithms = await get_jwks_client_and_algorithms()
    key = jwks_client.get_signing_key_from_jwt(token)
    token_version = jwt.decode(token, options={"verify_signature": False}).get("ver")
    # https://learn.microsoft.com/en-us/azure/active-directory/develop/access-tokens#token-formats
    # Upgrade to 2.0:
    # https://learn.microsoft.com/en-us/answers/questions/639834/how-to-get-access-token-version-20.html
    if token_version == "1.0":
        issuer = f"https://sts.windows.net/{settings.entraid_tenant_id}/"
        audience = f"api://{settings.entraid_client_id}"
    elif token_version == "2.0":
        issuer = f"https://login.microsoftonline.com/{settings.entraid_tenant_id}/v2.0"
        audience = settings.entraid_client_id
    else:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Auth error: Unsupported token version: {token_version}",
        )
    try:
        if settings.entraid_validate_issuer:
            claims = jwt.decode(
                jwt=token,
                key=key.key,
                algorithms=algorithms,
                audience=audience,
                issuer=issuer,
            )
        else:
            claims = jwt.decode(
                jwt=token,
                key=key.key,
                algorithms=algorithms,
                audience=audience,
            )
            # External users have their own tenant_id
            issuer_regex = r"https://login\.microsoftonline\.com/(common|organizations|consumers|[0-9a-fA-F-]{36})/v2\.0"
            if not re.match(issuer_regex, issuer):
                logger.debug(f"Couldn't match issuer for token: {token}")
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Auth error: Couldn't validate token",
                )
        logger.debug(claims)
    except Exception as e:
        logger.debug(f"Authentication error for token: {token}")
        logger.info(repr(e))
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Auth error: Couldn't validate token",
        )

    current_user = User(
        id=claims.get("oid"),
        name=claims.get("name"),
        email=claims.get("preferred_username"),
        roles=claims.get("roles", []),
    )
    logger.info(f"User authenticated: {current_user.name}")
    return current_user


def authorize(user: User, roles: list = None):
    if roles:
        if set(roles).issubset(user.roles):
            logger.info(f"User authorized: {user.name}")
            return user
        else:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Auth error: Missing roles for {user.name}: {', '.join(set(roles).difference(user.roles))}",
            )
    else:
        return user


async def authenticate(token: str = Header(default="", alias="Authorization")):
    """Dependency, returns a user object"""
    return await validate_token(token)


class Authorizer:
    """This class can be used to create dependencies that require specific roles:

    get_specific_user = Authorizer(roles=["role1", "role2"])

    Returns a user object.
    """

    def __init__(self, roles: list = None):
        self.roles = roles

    def __call__(self, current_user: User = Depends(authenticate)):
        return authorize(current_user, self.roles)


# Dependencies for RBAC
get_user = Authorizer()
get_admin = Authorizer(roles=["admin"])
