select cast(max(cast(coalesce(T0."U_JBC_CERTI",'0') as int))+1 as nvarchar(50))
from
OINV T0