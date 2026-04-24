from pydantic import BaseModel


class RateUpdate(BaseModel):
    usd_rate: float
    rmb_rate: float
