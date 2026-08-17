// Default mapping configuration promoted from the live browser export.
// This is the shipped default for new users. Do not hand-edit mappings.
export const defaultConfig = {
  "configVersion": 1,
  "applicationVersion": "1.0.0",
  "propertyMatch": {
    "type": "source",
    "sourceKey": "property|general|property name|1",
    "sourceSection": "General",
    "sourceField": "Property Name",
    "occurrence": 1
  },
  "unitMatch": {
    "type": "source",
    "sourceKey": "unit|general|property name|1",
    "sourceSection": "General",
    "sourceField": "Property Name",
    "occurrence": 1
  },
  "columnMappings": [
    {
      "targetColumnId": "property_reference",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|external id|1",
        "sourceSection": "General",
        "sourceField": "External ID",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "town",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|city|1",
        "sourceSection": "General",
        "sourceField": "City",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "address",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|property name|1",
        "sourceSection": "General",
        "sourceField": "Property Name",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "tenure_uk",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|tenure|1",
        "sourceSection": "General",
        "sourceField": "Tenure",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "area_sqft",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|occupied area|feb-2024|1",
        "sourceSection": "Occupied Area",
        "sourceField": "Feb-2024",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|building occupied area (sf)|feb-2024|1",
        "sourceSection": "Building Occupied Area (SF)",
        "sourceField": "Feb-2024",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "unit",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|suite|1",
        "sourceSection": "General",
        "sourceField": "Suite",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "tenant_name",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|tenant name|1",
        "sourceSection": "General",
        "sourceField": "Tenant Name",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "start_date",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|lease start|1",
        "sourceSection": "General",
        "sourceField": "Lease Start",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "review_date",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|variable|next review|1",
        "sourceSection": "Variable",
        "sourceField": "Next Review",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "expiry_date",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|lease expiry|1",
        "sourceSection": "General",
        "sourceField": "Lease Expiry",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "break_date",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|effective break|1",
        "sourceSection": "General",
        "sourceField": "Effective Break",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "mwault_year",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "contracted_rent_pa",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|contracted rent|1",
        "sourceSection": "General",
        "sourceField": "Contracted Rent",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|general|contracted rent|1",
        "sourceSection": "General",
        "sourceField": "Contracted Rent",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "contracted_rent_psf",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "deemed_rent_psf",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|net rent|1",
        "sourceSection": "General",
        "sourceField": "Net Rent",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|variable|current rent|1",
        "sourceSection": "Variable",
        "sourceField": "Current Rent",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "market_rent_pa",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|rental value (erv)|1",
        "sourceSection": "General",
        "sourceField": "Rental Value (ERV)",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|variable|rental value (erv)|1",
        "sourceSection": "Variable",
        "sourceField": "Rental Value (ERV)",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "market_rent_psf",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "current_void",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|variable|void|1",
        "sourceSection": "Variable",
        "sourceField": "Void",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "rent_free_month",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|variable|free rent (months)|1",
        "sourceSection": "Variable",
        "sourceField": "Free Rent (Months)",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "expiry_void_month",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "rent_free_expiry_month",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "cap_rate",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "source",
        "sourceKey": "unit|traditional valuation|hardcore yield|1",
        "sourceSection": "Traditional Valuation",
        "sourceField": "Hardcore Yield",
        "occurrence": 1
      },
      "ignore": false
    },
    {
      "targetColumnId": "niy",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|initial yield (at valuation date)|1",
        "sourceSection": "General",
        "sourceField": "Initial Yield (at Valuation Date)",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "ney",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|equivalent yield (at valuation date)|1",
        "sourceSection": "General",
        "sourceField": "Equivalent Yield (at Valuation Date)",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "ry",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|reversionary yield (at valuation date)|1",
        "sourceSection": "General",
        "sourceField": "Reversionary Yield (at Valuation Date)",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "cv_psf",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|say value|1",
        "sourceSection": "General",
        "sourceField": "Say Value",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "cap_ex",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "fair_value_april_2024",
      "propertyMapping": {
        "type": "source",
        "sourceKey": "property|general|say value|1",
        "sourceSection": "General",
        "sourceField": "Say Value",
        "occurrence": 1
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "fair_value_march_2026",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "pct_change_26_april_24",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "pct_change_x_ytd",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    },
    {
      "targetColumnId": "comments",
      "propertyMapping": {
        "type": "blank"
      },
      "unitMapping": {
        "type": "blank"
      },
      "ignore": false
    }
  ],
  "propertySummary": {
    "columnMappings": [
      {
        "targetColumnId": "property",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|property name|1",
          "sourceSection": "General",
          "sourceField": "Property Name",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "property_reference",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|external id|1",
          "sourceSection": "General",
          "sourceField": "External ID",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "contracted_rent_gross",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|contracted rent|1",
          "sourceSection": "General",
          "sourceField": "Contracted Rent",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "contracted_rent_net",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|net rent|1",
          "sourceSection": "General",
          "sourceField": "Net Rent",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "market_rent",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|rental value (erv)|1",
          "sourceSection": "General",
          "sourceField": "Rental Value (ERV)",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "niy",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|initial yield (at valuation date)|1",
          "sourceSection": "General",
          "sourceField": "Initial Yield (at Valuation Date)",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "ney",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|equivalent yield (at valuation date)|1",
          "sourceSection": "General",
          "sourceField": "Equivalent Yield (at Valuation Date)",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "ry",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|reversionary yield (at valuation date)|1",
          "sourceSection": "General",
          "sourceField": "Reversionary Yield (at Valuation Date)",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "market_value",
        "propertyMapping": {
          "type": "source",
          "sourceKey": "property|general|say value|1",
          "sourceSection": "General",
          "sourceField": "Say Value",
          "occurrence": 1
        },
        "ignore": false
      },
      {
        "targetColumnId": "market_value_change",
        "propertyMapping": {
          "type": "blank"
        },
        "ignore": false
      },
      {
        "targetColumnId": "market_rent_change",
        "propertyMapping": {
          "type": "blank"
        },
        "ignore": false
      },
      {
        "targetColumnId": "market_value_change_pct",
        "propertyMapping": {
          "type": "blank"
        },
        "ignore": false
      },
      {
        "targetColumnId": "draft_valuation_comments",
        "propertyMapping": {
          "type": "blank"
        },
        "ignore": false
      },
      {
        "targetColumnId": "draft_to_final_valuation_comments",
        "propertyMapping": {
          "type": "blank"
        },
        "ignore": false
      }
    ]
  }
};
