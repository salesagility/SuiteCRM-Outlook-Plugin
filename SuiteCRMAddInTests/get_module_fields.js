{"deleted": {
  "name": "deleted",
  "type": "bool",
  "group": "",
  "id_name": "",
  "label": "Deleted",
  "required": 0,
  "options": {
    "": {
      "name": "",
      "value": ""
    },
    "1": {
      "name": 1,
      "value": "Yes"
    },
    "2": {
      "name": 2,
      "value": "No"
    }
  },
  "related_module": "",
  "calculated": false,
  "len": ""
}}


{{
  "id": {
    "name": "id",
    "type": "id",
    "group": "",
    "id_name": "",
    "label": "ID",
    "required": 1,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "date_entered": {
    "name": "date_entered",
    "type": "datetime",
    "group": "",
    "id_name": "",
    "label": "Date Created:",
    "required": 1,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "date_modified": {
    "name": "date_modified",
    "type": "datetime",
    "group": "",
    "id_name": "",
    "label": "Date Modified",
    "required": 1,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "assigned_user_id": {
    "name": "assigned_user_id",
    "type": "assigned_user_name",
    "group": "",
    "id_name": "assigned_user_id",
    "label": "Assigned To:",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "assigned_user_name": {
    "name": "assigned_user_name",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "Assigned To:",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "modified_user_id": {
    "name": "modified_user_id",
    "type": "assigned_user_name",
    "group": "",
    "id_name": "modified_user_id",
    "label": "Modified By",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "created_by": {
    "name": "created_by",
    "type": "id",
    "group": "",
    "id_name": "",
    "label": "Created by",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": "36"
  },
  "deleted": {
    "name": "deleted",
    "type": "bool",
    "group": "",
    "id_name": "",
    "label": "Deleted",
    "required": 0,
    "options": {
      "": {
        "name": "",
        "value": ""
      },
      "1": {
        "name": 1,
        "value": "Yes"
      },
      "2": {
        "name": 2,
        "value": "No"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "from_addr_name": {
    "name": "from_addr_name",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "from_addr_name",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "reply_to_addr": {
    "name": "reply_to_addr",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "reply_to_addr",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "to_addrs_names": {
    "name": "to_addrs_names",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "to_addrs_names",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "cc_addrs_names": {
    "name": "cc_addrs_names",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "cc_addrs_names",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "bcc_addrs_names": {
    "name": "bcc_addrs_names",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "bcc_addrs_names",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "raw_source": {
    "name": "raw_source",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "raw_source",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "description_html": {
    "name": "description_html",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "description_html",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "description": {
    "name": "description",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "description",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "date_sent": {
    "name": "date_sent",
    "type": "datetime",
    "group": "",
    "id_name": "",
    "label": "Date Sent:",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "message_id": {
    "name": "message_id",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "Message ID",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": 255
  },
  "name": {
    "name": "name",
    "type": "name",
    "group": "",
    "id_name": "",
    "label": "Subject:",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": "255"
  },
  "type": {
    "name": "type",
    "type": "enum",
    "group": "",
    "id_name": "",
    "label": "Type",
    "required": 0,
    "options": {
      "out": {
        "name": "out",
        "value": "Sent"
      },
      "archived": {
        "name": "archived",
        "value": "Archived"
      },
      "draft": {
        "name": "draft",
        "value": "Draft"
      },
      "inbound": {
        "name": "inbound",
        "value": "Inbound"
      },
      "campaign": {
        "name": "campaign",
        "value": "Campaign"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": 100
  },
  "status": {
    "name": "status",
    "type": "enum",
    "group": "",
    "id_name": "",
    "label": "Email Status:",
    "required": 0,
    "options": {
      "archived": {
        "name": "archived",
        "value": "Archived"
      },
      "closed": {
        "name": "closed",
        "value": "Closed"
      },
      "draft": {
        "name": "draft",
        "value": "In Draft"
      },
      "read": {
        "name": "read",
        "value": "Read"
      },
      "replied": {
        "name": "replied",
        "value": "Replied"
      },
      "sent": {
        "name": "sent",
        "value": "Sent"
      },
      "send_error": {
        "name": "send_error",
        "value": "Send Error"
      },
      "unread": {
        "name": "unread",
        "value": "Unread"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": 100
  },
  "flagged": {
    "name": "flagged",
    "type": "bool",
    "group": "",
    "id_name": "",
    "label": "Flagged:",
    "required": 0,
    "options": {
      "": {
        "name": "",
        "value": ""
      },
      "1": {
        "name": 1,
        "value": "Yes"
      },
      "2": {
        "name": 2,
        "value": "No"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "reply_to_status": {
    "name": "reply_to_status",
    "type": "bool",
    "group": "",
    "id_name": "",
    "label": "Reply To Status:",
    "required": 0,
    "options": {
      "": {
        "name": "",
        "value": ""
      },
      "1": {
        "name": 1,
        "value": "Yes"
      },
      "2": {
        "name": 2,
        "value": "No"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "intent": {
    "name": "intent",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "Intent",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": 100,
    "default_value": "pick"
  },
  "mailbox_id": {
    "name": "mailbox_id",
    "type": "id",
    "group": "",
    "id_name": "",
    "label": "LBL_MAILBOX_ID",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": "36"
  },
  "parent_name": {
    "name": "parent_name",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "parent_name",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "parent_type": {
    "name": "parent_type",
    "type": "varchar",
    "group": "",
    "id_name": "",
    "label": "parent_type",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": 100
  },
  "parent_id": {
    "name": "parent_id",
    "type": "id",
    "group": "",
    "id_name": "",
    "label": "parent_id",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": "36"
  },
  "category_id": {
    "name": "category_id",
    "type": "enum",
    "group": "",
    "id_name": "",
    "label": "Category",
    "required": 0,
    "options": {
      "": {
        "name": "",
        "value": ""
      },
      "Archived": {
        "name": "Archived",
        "value": "Archived"
      },
      "Sales": {
        "name": "Sales",
        "value": "Sales"
      },
      "Marketing": {
        "name": "Marketing",
        "value": "Marketing"
      },
      "FeedBack": {
        "name": "FeedBack",
        "value": "Feedback"
      }
    },
    "related_module": "",
    "calculated": false,
    "len": 100
  },
  "modified_by_name": {
    "name": "modified_by_name",
    "type": "assigned_user_name",
    "group": "",
    "id_name": "modified_user_id",
    "label": "Modified By",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": ""
  },
  "created_by_name": {
    "name": "created_by_name",
    "type": "id",
    "group": "",
    "id_name": "",
    "label": "Created by",
    "required": 0,
    "options": [],
    "related_module": "",
    "calculated": false,
    "len": "36"
  }
}}