/* MDB Tools - A library for reading MS Access database file
 * Copyright (C) 2000 Brian Bruns
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Library General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Library General Public License for more details.
 *
 * You should have received a copy of the GNU Library General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */

#include "mdbtools.h"

/* Read an integer value from a column's current value in the page buffer,
 * handling BYTE, INT, and LONGINT column types without a string round-trip. */
static long read_int_col(MdbHandle *mdb, MdbColumn *col)
{
	if (!col->cur_value_len)
		return 0;
	switch (col->col_type) {
		case MDB_BYTE:
			return mdb_get_byte(mdb->pg_buf, col->cur_value_start);
		case MDB_INT:
			/* Cast to short first to sign-extend the 16-bit value before widening to long */
			return (long)(short)mdb_get_int16(mdb->pg_buf, col->cur_value_start);
		case MDB_LONGINT:
		default:
			return mdb_get_int32(mdb->pg_buf, col->cur_value_start);
	}
}

const char *
mdb_get_objtype_string(int obj_type)
{
    static const char *type_name[] = {"Form",
			"Table",
			"Macro",
			"System Table",
			"Report",
			"Query",
			"Linked Table",
			"Module",
			"Relationship",
			"Unknown 0x09",
			"User Info",
			"Database"
		};

	if (obj_type >= (int)(sizeof(type_name)/sizeof(type_name[0]))) {
		return NULL;
	} else {
		return type_name[obj_type]; 
	}
}

void mdb_free_catalog(MdbHandle *mdb)
{
	guint i, j;
	MdbCatalogEntry *entry;

	if ((!mdb) || (!mdb->catalog)) return;
	for (i=0; i<mdb->catalog->len; i++) {
		entry = (MdbCatalogEntry *)g_ptr_array_index(mdb->catalog, i);
		if (entry) {
			if (entry->props) {
				for (j=0; j<entry->props->len; j++)
					mdb_free_props(g_ptr_array_index(entry->props, j));
				g_ptr_array_free(entry->props, TRUE);
			}
			g_free(entry);
		}
	}
	g_ptr_array_free(mdb->catalog, TRUE);
	mdb->catalog = NULL;
}

GPtrArray *mdb_read_catalog (MdbHandle *mdb, int objtype)
{
	MdbCatalogEntry *entry, msysobj;
	MdbTableDef *table;
	char *obj_name = NULL;
	char *obj_props = NULL;
	int type;
	int i;
	MdbColumn *col_id, *col_type, *col_flags, *col_props;
	int col_id_num, col_type_num, col_flags_num;
	int kkd_size_ole;

	if (!mdb) return NULL;
	if (mdb->catalog) mdb_free_catalog(mdb);
	mdb->catalog = g_ptr_array_new();
	mdb->num_catalog = 0;

	obj_name = malloc(mdb->bind_size);
	obj_props = malloc(mdb->bind_size);

	/* dummy up a catalog entry so we may read the table def */
	memset(&msysobj, 0, sizeof(MdbCatalogEntry));
	msysobj.mdb = mdb;
	msysobj.object_type = MDB_TABLE;
	msysobj.table_pg = 2;
	snprintf(msysobj.object_name, sizeof(msysobj.object_name), "%s", "MSysObjects");

	/* mdb_table_dump(&msysobj); */

	table = mdb_read_table(&msysobj);
    if (!table) {
        fprintf(stderr, "Unable to read table %s\n", msysobj.object_name);
        mdb_free_catalog(mdb);
        goto cleanup;
    }

	if (!mdb_read_columns(table)) {
		fprintf(stderr, "Unable to read columns of table %s\n", msysobj.object_name);
		mdb_free_catalog(mdb);
		goto cleanup;
	}

	/* Bind integer columns with NULL so their cur_value_start is tracked
	 * without the overhead of converting to string and back. */
    if ((col_id_num = mdb_bind_column_by_name(table, "Id", NULL, NULL)) == -1 ||
        mdb_bind_column_by_name(table, "Name", obj_name, NULL) == -1 ||
        (col_type_num = mdb_bind_column_by_name(table, "Type", NULL, NULL)) == -1 ||
        (col_flags_num = mdb_bind_column_by_name(table, "Flags", NULL, NULL)) == -1) {
        fprintf(stderr, "Unable to bind columns from table %s (%d columns found)\n",
                msysobj.object_name, table->num_cols);
        mdb_free_catalog(mdb);
        goto cleanup;
    }
    if ((i = mdb_bind_column_by_name(table, "LvProp", obj_props, &kkd_size_ole)) == -1) {
        fprintf(stderr, "Unable to bind column %s from table %s\n", "LvProp", msysobj.object_name);
        mdb_free_catalog(mdb);
        goto cleanup;
    }
	col_id    = g_ptr_array_index(table->columns, col_id_num - 1);
	col_type  = g_ptr_array_index(table->columns, col_type_num - 1);
	col_flags = g_ptr_array_index(table->columns, col_flags_num - 1);
	col_props = g_ptr_array_index(table->columns, i-1);

	mdb_rewind_table(table);

	while (mdb_fetch_row(table)) {
		type = (int)read_int_col(mdb, col_type);
		if (objtype==MDB_ANY || type == objtype) {
			//fprintf(stderr, "obj_id: %10ld objtype: %-3d (0x%04x) obj_name: %s\n",
			//	(read_int_col(mdb, col_id) & 0x00FFFFFF), type, type, obj_name);
			entry = g_malloc0(sizeof(MdbCatalogEntry));
			entry->mdb = mdb;
			snprintf(entry->object_name, sizeof(entry->object_name), "%s", obj_name);
			entry->object_type = (type & 0x7F);
			entry->table_pg = (unsigned long)read_int_col(mdb, col_id) & 0x00FFFFFF;
			entry->flags = read_int_col(mdb, col_flags);
			mdb->num_catalog++;
			g_ptr_array_add(mdb->catalog, entry);
			if (kkd_size_ole) {
				size_t kkd_len;
				void *kkd = mdb_ole_read_full(mdb, col_props, &kkd_len);
				//mdb_buffer_dump(kkd, 0, kkd_len);
				if (kkd) {
					entry->props = mdb_kkd_to_props(mdb, kkd, kkd_len);
					free(kkd);
				}
			}
		}
	}
	//mdb_dump_catalog(mdb, MDB_TABLE);
 
cleanup:
	if (table)
		mdb_free_tabledef(table);

	free(obj_name);
	free(obj_props);

    return mdb->catalog;
}


MdbCatalogEntry *
mdb_get_catalogentry_by_name(MdbHandle *mdb, const gchar* name)
{
	unsigned int i;
	MdbCatalogEntry *entry;

	for (i=0; i<mdb->num_catalog; i++) {
		entry = g_ptr_array_index(mdb->catalog, i);
		if (!g_ascii_strcasecmp(entry->object_name, name))
			return entry;
	}
	return NULL;
}

void 
mdb_dump_catalog(MdbHandle *mdb, int obj_type)
{
	unsigned int i;
	MdbCatalogEntry *entry;

	mdb_read_catalog(mdb, obj_type);
	for (i=0;i<mdb->num_catalog;i++) {
                entry = g_ptr_array_index(mdb->catalog,i);
		if (obj_type==MDB_ANY || entry->object_type==obj_type) {
			printf("Type: %-12s Name: %-48s Page: %06lx\n",
			mdb_get_objtype_string(entry->object_type) ?: "Unknown",
			entry->object_name,
			entry->table_pg);
		}
	}
	return;
}

