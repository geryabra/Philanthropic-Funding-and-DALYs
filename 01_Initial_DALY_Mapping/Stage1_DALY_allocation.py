"""Run after Org_mission_to_DALY_v8.py. Place this script beside its inputs.
Inputs: FINAL_DATASET_STAGE1.xlsx (final_dataset, full_enriched_debug),
DALYS_2022.xlsx (Dalys), Org_mission_to_DALY_v8.py (dictionary definitions),
Other_Neoplasms_Mission_Review_2022.xlsx (Review_390).
Output: Stage1_DALY_allocation_2022.xlsx (final_dataset, before_after, allocation).
Revenue is apportioned in cents. Pending cases retain their original revenue.
"""
import ast
import re
from pathlib import Path
from decimal import Decimal
import pandas as pd

BASE = Path(__file__).resolve().parent

def norm(value):
    return ' '.join(re.findall(r'[a-z0-9]+', str(value).lower()))

def main():
    daly = pd.read_excel(BASE/'DALYS_2022.xlsx', sheet_name='Dalys')
    data = pd.read_excel(BASE/'FINAL_DATASET_STAGE1.xlsx', sheet_name='full_enriched_debug')
    original = pd.read_excel(BASE/'FINAL_DATASET_STAGE1.xlsx', sheet_name='final_dataset')
    reviews = pd.read_excel(BASE/'Other_Neoplasms_Mission_Review_2022.xlsx', sheet_name='Review_390', dtype={'EIN':str})
    reviews = reviews.set_index('EIN',verify_integrity=True)
    if not original[['Revenue','Org_mission','NTEE']].equals(data[['Revenue','Org_mission','NTEE']]):
        raise ValueError('Stage 1 sheets have inconsistent row order')
    if daly.Disease.duplicated().any() or not daly.DALYs.gt(0).all():
        raise ValueError('Invalid disease-level DALYs')
    weights = dict(zip(daly.Disease,daly.DALYs.map(lambda x:Decimal(str(x)))))
    tree = ast.parse((BASE/'Org_mission_to_DALY_v8.py').read_text(encoding='utf-8-sig'))
    dictionaries = {n.targets[0].id:ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign)
                    and isinstance(n.targets[0],ast.Name) and n.targets[0].id in ('manual_aliases','disease_merge_aliases')}
    lookup = {}
    for col in ['Level_2','Level_3','Disease']:
        for label,g in daly.groupby(col):
            lookup[norm(label)] = set(g.Disease)
    aliases = {norm(k):norm(v) for k,v in dictionaries['disease_merge_aliases'].items()}
    def resolve(label):
        key = norm(label)
        return lookup.get(aliases.get(key,key),set())
    terms = {norm(x):x for x in daly.Disease if not norm(x).startswith('other ')}
    terms.update({norm(k):v for k,v in dictionaries['manual_aliases'].items()})
    terms.update({'ovarian cancer':'Ovary cancer','ovaraian cancer':'Ovary cancer','lung cancer':'Trachea, bronchus, lung cancers',
                  'pancreatic cancer':'Pancreas cancer','brain cancer':'Brain and nervous system cancers',
                  'colon cancer':'Colon and rectum cancers','colorectal cancer':'Colon and rectum cancers',
                  'glioblastoma':'Brain and nervous system cancers','gliomas':'Brain and nervous system cancers'})
    fields = {norm(k):k for k in daly.Level_2 if not norm(k).startswith('other ')}
    fields.update({'cancer':'Malignant neoplasms','cancers':'Malignant neoplasms','oncology':'Malignant neoplasms',
                   'cardiac':'Cardiovascular diseases','cardiology':'Cardiovascular diseases','cardiovascular':'Cardiovascular diseases','heart disease':'Cardiovascular diseases',
                   'neurology':'Neurological conditions','dermatology':'Skin diseases','dentistry':'Oral conditions',
                   'gastroenterology':'Digestive diseases','pulmonology':'Respiratory diseases','nephrology':'Kidney diseases'})
    for k in fields: terms.pop(k,None)
    def scan(text):
        padded = ' '+norm(text)+' '
        specific, broad = {},{}
        for phrase, pair in [('breast or ovarian cancer',('Breast cancer','Ovary cancer')),
                             ('breast and ovarian cancer',('Breast cancer','Ovary cancer')),
                             ('breast and cervical cancer',('Breast cancer','Cervix uteri cancer'))]:
            if ' '+phrase+' ' in padded:
                specific[phrase]=set(pair)
        for mapping,found in [(terms,specific),(fields,broad)]:
            for k,v in mapping.items():
                if ' '+k+' ' in padded:
                    found[k]=resolve(v) if mapping is terms else lookup[norm(v)]
        return specific,broad
    def union(matches):
        return set().union(*matches.values())
    records=[]; allocations=[]
    for row in data.itertuples():
        ein=str(int(row.extract_ein)).zfill(9)
        name=row.extract_name
        mission='' if pd.isna(row.Org_mission) else row.Org_mission
        sp,br=scan(mission); ns,nb=scan(name)
        prior='' if pd.isna(row.Disease_norm_daly) else str(row.Disease_norm_daly)
        targets=set(); source=''; status=''; evidence=''
        review=reviews.loc[ein] if ein in reviews.index else None
        if review is not None and review.Org_mission!=mission:
            raise ValueError('Review mission changed: '+ein)
        if review is not None and review.Proposed_decision in ['Один тип','Несколько типов']:
            targets=set(review.DALY_target.split('; '));source='Mission review';evidence=review.Mission_quote
        elif review is not None and review.Proposed_decision!='Общий рак':
            status='Review: '+review.Proposed_decision; evidence=review.Mission_quote;source='Mission review'
            if ns or nb: evidence+=' | Name: '+', '.join([*ns,*nb])
        else:
            # Retain previously recognized diagnoses, including residual crosswalks.
            saved=[] if pd.isna(row.matched_specific_terms) else row.matched_specific_terms.split('; ')
            if review is not None:
                sp={};br={'cancer':lookup['malignant neoplasms']};saved=[]
            existing=set()
            for label in saved:
                group=resolve(label)
                if norm(label) in ('other neoplasms','other circulatory diseases','other oral disorders') and br and group <= union(br):
                    continue
                existing.update(group)
                if not group: status='Review: unresolved Stage 1 diagnosis'
            targets=existing|union(sp)
            # A broad alias nested inside an explicit specific diagnosis adds no extra leaves.
            if targets:
                if br and any(not (g & targets) for g in br.values()):
                    status='Review: multiple scopes'
                source='Mission'; evidence=', '.join([*sp,*br]) or '; '.join(saved)
            elif br:
                targets=union(br);source='Mission field';evidence=', '.join(br)
                if ns:
                    if union(ns) and union(ns)<=targets and all(ns.values()):
                        targets=union(ns);source='Mission field + name';evidence+=' | '+', '.join(ns)
                    else:status='Review: conflicting or unresolved name'
            elif ns or nb:
                targets=union(ns) or union(nb);source='Organization name';evidence=', '.join([*ns,*nb])
                if any(not x for x in ns.values()) or (ns and nb and not union(ns)<=union(nb)):
                    status='Review: ambiguous name'
            if ns and targets and union(ns)-targets:status='Review: conflicting name'
            if any(not x for x in sp.values()):status='Review: unresolved term'
            if (source=='Organization name' and any(k in ns for k in ['trauma','falls','als','aids','blind'])):
                status='Review: ambiguous name term'
            if source in ('Mission field','Organization name','Mission field + name') and re.search(
                    r'\b(veterinary|canine|pediatric|childhood|rare cancer|sarcoma|neuroblastoma|neuroendocrine|gynecologic|retinoblastoma)\b',
                    norm(name)+' '+norm(mission)):
                status='Review: restricted scope'
            if source=='Mission' and len(saved)>1 and any(k in sp for k in ['trauma','falls']):
                status='Review: contextual mention'
        if not targets and not status:status='No mapped evidence'
        if targets and not status:status='Mapped' if pd.notna(row.Revenue) else 'Missing revenue'
        if status=='Mapped':
            labels=sorted(targets);den=sum(weights[t] for t in labels)
            cents=int(Decimal(str(row.Revenue))*100)
            shares=[abs(cents)*weights[t]/den for t in labels]
            parts=[int(x) for x in shares]
            for i in sorted(range(len(labels)),key=lambda i:shares[i]-parts[i],reverse=True)[:abs(cents)-sum(parts)]:
                parts[i]+=1
            assert sum(parts)==abs(cents)
            for t,part in zip(labels,parts):
                allocations.append([ein,name,t,float(weights[t]/den),part*(1 if cents>=0 else -1)/100,source])
        records.append([ein,name,prior,'; '.join(sorted(targets)),status,source,evidence])
    audit=pd.DataFrame(records,columns=['EIN','Organization_Name','Baseline_DALY','Candidate_DALYs','Allocation_status','Evidence_source','Evidence'])
    assert audit.EIN.is_unique and len(audit)==len(original)
    final=pd.concat([audit.iloc[:,:2],original,audit.iloc[:,2:]],axis=1)
    allocation=pd.DataFrame(allocations,columns=['EIN','Organization_Name','Disease','Weight','Allocated_Revenue','Evidence_source'])
    changes=final.loc[(final.Candidate_DALYs.map(norm)!=final.Baseline_DALY.map(norm))|final.Allocation_status.str.startswith('Review')]
    with pd.ExcelWriter(BASE/'Stage1_DALY_allocation_2022.xlsx') as writer:
        for sheet,frame in [('final_dataset',final),('before_after',changes),('allocation',allocation)]:
            frame.to_excel(writer,sheet_name=sheet,index=False,freeze_panes=(1,2))
    print(final.Allocation_status.value_counts(),final.Evidence_source.value_counts())

if __name__=='__main__':
    main()
