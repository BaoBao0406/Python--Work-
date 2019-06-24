#! python3
# stripJunkSimpleSalesforce - Convert the result from Salesforce data to a readable data.
# Copy from GitHib - https://github.com/simple-salesforce/simple-salesforce/issues/283

def stripJunkSimpleSalesforce(structure):
    from collections import OrderedDict
    import json
    def fixdict(struc, keyChainList, parentListRecord, parentListAddColsAndVals, parentListKillCols, passingContext):
        if 'records' in struc.keys() and type(struc['records']) is list:
            parentListRecord[keyChainList[0]] = struc['records']
            fixlist(parentListRecord[keyChainList[0]], 'fixDict')
        else:
            if 'attributes' in struc.keys():
                struc.pop('attributes')
            for k in struc.keys():
                if k != 'attributes':
                    if type(struc[k]) in [dict, OrderedDict]:
                        parentListAddColsAndVals, parentListKillCols = fixdict(struc[k], keyChainList + [k], parentListRecord, parentListAddColsAndVals, parentListKillCols, 'fixDict')
                    if type(struc[k]) not in [None, list, dict, OrderedDict] and len(keyChainList) > 0:
                        parentListAddColsAndVals['.'.join(keyChainList+[k])] = struc[k]
                        if passingContext == 'fixDict':
                            parentListKillCols.add(keyChainList[0])
        return parentListAddColsAndVals, parentListKillCols
    def fixlist(struc, passingContext):
        if type(struc) is list and len(struc) > 0:
            for x in struc:
                if type(x) in [dict, OrderedDict]:
                    listAddColsAndVals, listKillCols = fixdict(x, [], x, {}, set(), 'fixList')
                    if len(listAddColsAndVals) > 0:
                        for k,v in listAddColsAndVals.items():
                                x[k] = v
                    if len(listKillCols) > 0:
                        for k in listKillCols:
                            if k in x.keys():
                                x.pop(k)
        return
    outerStructure = None
    # I can't get this working for Ordered Dicts, so giving up and having it return a plain dict.  Leaving some junk from earlier attempt behind.
    if type(structure) in [dict, OrderedDict] and 'records' in structure.keys() and type(structure['records']) is list:
        if type(structure) is OrderedDict:
            outerStructure = json.loads(json.dumps(structure))['records']
        else:
            outerStructure = structure['records']
    if type(outerStructure) is list:
        fixlist(outerStructure, 'outermost')
    return outerStructure