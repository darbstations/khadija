# -*- coding: utf-8 -*-
MONTHS = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]

def ytd(agg, values):
    nums = [v for v in values if v is not None]
    if not nums: return None
    if agg == "SUM":  return sum(nums)
    if agg == "AVG":  return sum(nums)/len(nums)
    return nums[-1]  # LAST

def achievement(target, polarity, ytd_val):
    if target is None or ytd_val is None: return None
    if target == 0: return 1.0 if ytd_val <= 0 else 0.0
    if polarity == "↓": return (target/ytd_val) if ytd_val else 0.0
    return ytd_val/target if target else None

def status(ach):
    if ach is None: return "—"
    if ach >= 1: return "✅ محقق"
    if ach >= 0.85: return "🟡 قريب"
    return "🔴 تحت الهدف"

def status_class(ach):
    if ach is None: return "muted"
    if ach >= 1: return "ok"
    if ach >= 0.85: return "warn"
    return "bad"

def forecast(agg, ytd_val, n_months):
    if ytd_val is None: return None
    if agg == "SUM":
        if not n_months: return None
        return ytd_val / n_months * 12
    return ytd_val

def fmt_value(v, fmt):
    if v is None: return "—"
    if fmt == "pct":  return f"{v*100:.0f}%"
    if fmt == "int":  return f"{v:,.0f}"
    if fmt == "num1": return f"{v:,.1f}"
    if fmt == "rial": return f"{v:,.0f} ر.س"
    return str(v)
