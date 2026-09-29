"""Builds the synthetic Banesco savings statement as text in three layouts (columns, one line, one cell per line).
Invented holder, rows and amounts; every balance follows from the one before, and the totals match the rows.
Run: python3 make_statement_text.py && python3 make_synthetic_pdf.py"""
HOLDER = "ANA MARIA PEREZ SOTO"
opening = 50000.00
rows = [("03/08/2026", "Lbtr Pedro Gomez Ramirez", 1000.00, 'D'), ("03/08/2026", "Imp. Art. 12 Ley 288-04", 2.00, 'D'),
        ("05/08/2026", "Transf Recibida Maria Lopez Rosa", 1000.00, 'C'),       # another description: a third party
        ("10/08/2026", "Lbtr Ana Maria Perez Sot", 5000.00, 'D'), ("17/08/2026", "Ach Ibanking", 6500.00, 'C'),
        ("20/08/2026", "Transf Recibida Ana Maria Perez Soto", 700.00, 'C'),    # another description naming the holder
        ("24/08/2026", "Ach Ibanking", 2300.00, 'C'), ("26/08/2026", "Pago Renta Apartamento 5", 20000.00, 'D'),
        ("28/08/2026", "Lbtr Ana Maria Perez Sot", 3000.00, 'C'), ("31/08/2026", "Intereses", 150.25, 'C'),
        ("31/08/2026", "Lbtr Luis Gomez Diaz", 1200.00, 'C'), ("31/08/2026", "Retencion De Intereses", 15.03, 'D')]
bal, out = opening, []
for d, desc, amt, kind in rows:
    bal = round(bal + (amt if kind == 'C' else -amt), 2); out.append((d, desc, amt, kind, bal))
credits = sum(a for _, _, a, k, _ in out if k == 'C'); debits = sum(a for _, _, a, k, _ in out if k == 'D')
f = lambda x: f"{x:,.2f}"
header = ["Tus Finanzas", "Estado de Cuenta de Ahorros", "Agosto 2026", HOLDER, "Cuenta regulatoria: DO00BANS00000000000000000000",
  "Tipo de cuenta: Cuenta De Ahorros                                          Estado: ACTIVA",
  "Fecha de apertura:              15/JUN/2026   Moneda:                                                           DOP",
  "N° de cuenta:                  00000000000    Fecha al corte:                                          31/AGO/2026", "Resumen de tu cuenta",
  "                                                 RD$                                                    RD$",
  f"Balance mes anterior:                             {f(opening):>12}         Débitos del mes:                                       {f(debits):>12}",
  f"Balance al corte:                                 {f(bal):>12}         Créditos del mes:                                      {f(credits):>12}",
  "Impuestos del mes:                                      2.00         Intereses acumulados al corte:                            150.25",
  "Detalle de tus transacciones", HOLDER, "Transacciones en RD$ Pesos Dominicanos"]
table_head = ["           Fecha", "        Transacción                Descripción de Movimientos                         Débito               Crédito              Balance"]
def line(d, desc, amt, kind, b):
    return f"         {d} {desc:<44}" + (f"{f(amt):>20}{'':>21}" if kind == 'D' else f"{'':>20}{f(amt):>21}") + f"{f(b):>21}"
layout = header + table_head + [line(*r) for r in out[:7]] + ["\f"] + table_head + [line(*r) for r in out[7:]] + ["Página 2 de 2"]
open('banesco_savings_layout.txt', 'w').write("\n".join(layout) + "\n")
open('banesco_savings_flat.txt', 'w').write(" ".join(" ".join(layout).split()) + "\n")
cells = [l.strip() for l in header + table_head]
for d, desc, amt, kind, b in out: cells += [d, desc, f(amt), f(b)]
open('banesco_savings_cells.txt', 'w').write("\n".join(c for c in cells if c) + "\n")
print("credits", f(credits), "debits", f(debits), "closing", f(bal))
