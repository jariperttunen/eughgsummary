## Columns in CRFReporter excel files
## Used in eululucftotal.y
netco2_col=1
ch4_col=2
n2o_col=3

def sum_emissions(e1,e2):
    """Utility function to sum emissions. Emission can be a number or  anotation key"""
    try:#Both are numbers
        return float(e1)+float(e2)
    except:#The other one is not 
        try:#Check first e1
            return float(e1)
        except:
            try:#Check e2
                return float(e2)
            except:#Both are notation keys
                return e1+','+e2

def convertco2eq(e,gwp):
    """Convert emission 'e' to CO2eq, 'e' can be notation key"""
    try:
        return float(gwp*e)
    except:
        return e
