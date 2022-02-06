import os

from fastapi import FastAPI, HTTPException, Security, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security.api_key import APIKeyHeader


async def validate_api_key(api_key: str = Security(APIKeyHeader(name="Authorization"))):
    """Validate the API_KEY as delivered by the Authorization header
    It is recommended to always set a unique XLWINGS_API_KEY as environment variable.
    Without an env var, it expects "DEVELOPMENT" as the API_KEY, which is insecure.
    """
    if api_key != os.getenv("XLWINGS_API_KEY", "DEVELOPMENT"):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid API Key",
        )


# Require the API_KEY for every endpoint
app = FastAPI(dependencies=[Security(validate_api_key)])

# Excel on the web and our Python backend are on different origins, so we'll need to
# enable CORS (Google Sheets doesn't doesn't use CORS and will ingore this)
app.add_middleware(
    CORSMiddleware,
    allow_origin_regex=r"https://.*.officescripts.microsoftusercontent.com",
    allow_methods=["POST"],
    allow_headers=["*"],
)
