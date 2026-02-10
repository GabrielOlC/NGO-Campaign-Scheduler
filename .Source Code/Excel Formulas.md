### tbASchedule: _Mensagem

> Purpose: Allow users to write dynamic messages without the need to use excel formulas, but simple tags (e.g., `<name>`).
> 
> Context: User writes the message using the tags, the code fills the tags with the proper text and on right-click the msg is copied to the user clipboard through VBA

```excel-formula
=IF([@Mensagem]<>"",
  LET(
    vMsg,[@Mensagem],
    vName, TRIM([@[Nome / Tel ]]),
    vDate,TEXT([@Data],"dd/mm/yy"),
    vHour,TEXT([@Data],"HH:MM"),
    vToken,
      LET(
        lmdTb,LAMBDA(lm,
          CHOOSE(lm,tbDBTokens[Senha],tbDBTokens[FK_IDAgendamento], tbDBTokens[Status])
        ),
        TEXTJOIN(", ",FALSE, FILTER(lmdTb(1),(lmdTb(2)=[@ID])*(lmdTb(3)<>"Cancelado"),""))
      ),

    rem,N("Tags to substitute"),
    vSa,SUBSTITUTE(vMsg,"<nome>",vName),
    vSb,SUBSTITUTE(vSa,"<data>",vDate),
    vSc,SUBSTITUTE(vSb,"<hora>",vHour),
    vSd,SUBSTITUTE(vSc,"<senhas>",vToken),
    vSd
  ),""
),
    vSd
  ),""
)
```

### tbASchedule: Senhas

> Purpose: Summarize to the user the current state of the local database
> 
> Context: When user double-click VBA updates the database creating or cancelling tokens, right-click open forms for token transfer.

```excel-formula
=LET(
  vTb, HSTACK(tbDBTokens[Senha], tbDBTokens[FK_IDAgendamento], tbDBTokens[Tipo], tbDBTokens[Status]),
  vID, [@ID],
  vFiltered, FILTER(vTb, (INDEX(vTb,,2)=vID)*(INDEX(vTb,,4)<>"Cancelado"),""),

  v,IF(@vFiltered="","",
    LET(
      lmdTbCol, LAMBDA(lm,
        INDEX(vFiltered,,lm)
      ),
      vTypes, {"CF","CM","FF","FM"},
      vCounts, MAP(vTypes,
        LAMBDA(lm,
          SUM(--(lmdTbCol(3)=lm))
        )
      ),

      vPairs, MAP(vTypes, vCounts,
        LAMBDA(lmTypes,lmCounts,
          IF(lmCounts>0, lmTypes & "(" & lmCounts & ")", ""))
      ),
      v, TEXTJOIN(", ", TRUE, vPairs),
      v
    )
  ),
  v
)
```

### tbDBTokens: Senhas

> Purpose: Allows user to change the tokens format at will
> 
> Context: 

```excel-formula
=[@Tipo] & sys_cnYearCode & "." & DEC2HEX([@ID])
```
