//------------------------------------------------------------------------------
// Project:  Shapelib
// Purpose:  Sample application for dumping contents of a shapefile to
//           the terminal in human readable form.
// Author:   Frank Warmerdam, warmerda@home.com
//           Pascal translation Alexander Weidauer, alex.weidauer@huckfinn.de
//------------------------------------------------------------------------------
// Copyright (c) 1999, Frank Warmerdam
//
// This software is available under the following "MIT Style" license,
// or at the option of the licensee under the LGPL (see LICENSE.LGPL).  This
// option is discussed in more detail in shapelib.html.
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
// OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// aDEALINGS IN THE SOFTWARE.
//------------------------------------------------------------------------------
Unit ShpAPI129;
Interface

Const

// ----------------------------------------------------------------------------
// Current DLL release
// ----------------------------------------------------------------------------
 LibName = 'shapelib129.dll';

// ----------------------------------------------------------------------------
// Configuration options.
// ----------------------------------------------------------------------------
//  Should the DBFReadStringAttribute() strip leading and trailing white space?                                            }
// ----------------------------------------------------------------------------
{$DEFINE TRIM_DBF_WHITESPACE}

// ----------------------------------------------------------------------------
// Should we write measure values to the Multipatch object?
// Reportedly ArcView crashes if we do write it, so for now it is disabled.                                                     }
// ----------------------------------------------------------------------------
{$DEFINE DISABLE_MULTIPATCH_MEASURE}

Type
// ----------------------------------------------------------------------------
// Basic definitions
// ----------------------------------------------------------------------------
  PDouble = ^Double;
  PLongInt = ^LongInt;
  PLongIntArray = ^TLongIntArray;
  TLongIntArray = Array Of LongInt;
  PDoubleArray = ^TDoubleArray;
  TDoubleArray = Array  Of Double;

// ----------------------------------------------------------------------------
// Shape info record 
// ----------------------------------------------------------------------------
  PSHPInfo = ^SHPInfo;
  SHPInfo = Record
    fpSHP: ^File;
    fpSHX: ^File;
    nShapeType: LongInt;
    nFileSize: LongInt;
    nRecords: LongInt;
    nMaxRecords: LongInt;
    panRecOffset: PLongInt;
    panRecSize: PLongInt;
    adBoundsMin: Array[0..3] Of Double;
    adBoundsMax: Array[0..3] Of Double;
    bUpdated: LongInt;
  End;

  SHPHandle = PSHPInfo;

// ----------------------------------------------------------------------------
// Shape types (nSHPType)
// ----------------------------------------------------------------------------
Const
  SHPT_NULL = 0;
  SHPT_POINT = 1;
  SHPT_ARC = 3;
  SHPT_POLYGON = 5;
  SHPT_MULTIPOINT = 8;
  SHPT_POINTZ = 11;
  SHPT_ARCZ = 13;
  SHPT_POLYGONZ = 15;
  SHPT_MULTIPOINTZ = 18;
  SHPT_POINTM = 21;
  SHPT_ARCM = 23;
  SHPT_POLYGONM = 25;
  SHPT_MULTIPOINTM = 28;
  SHPT_MULTIPATCH = 31;
// ----------------------------------------------------------------------------
// Part types - everything but SHPT_MULTIPATCH just uses SHPP_RING.                                                       }
// ----------------------------------------------------------------------------
  SHPP_TRISTRIP = 0;
  SHPP_TRIFAN = 1;
  SHPP_OUTERRING = 2;
  SHPP_INNERRING = 3;
  SHPP_FIRSTRING = 4;
  SHPP_RING = 5;

// ----------------------------------------------------------------------------
// SHPObject - represents on shape (without attributes) read from the .shp file.                                              }
//             -1 is unknown/unassigned  
// ----------------------------------------------------------------------------
Type
  PPShpObject = ^PShpObject;
  PShpObject = ^ShpObject;
  ShpObject = Record
    nSHPType: LongInt;
    nShapeId: LongInt;
    nParts: LongInt;
    panPartStart: TLongIntArray;
    panPartType: TLongIntArray;
    nVertices: LongInt;
    padfX:  TDoubleArray;
    padfY:  TDoubleArray;
    padfZ:  TDoubleArray;
    padfM:  TDoubleArray;
    dfXMin: Double;
    dfYMin: Double;
    dfZMin: Double;
    dfMMin: Double;
    dfXMax: Double;
    dfYMax: Double;
    dfZMax: Double;
    dfMMax: Double;
  End;
// ----------------------------------------------------------------------------
// SHP API Prototypes                                               
// ----------------------------------------------------------------------------
Function SHPOpen (pszShapeFile: Pchar; pszAccess: Pchar ) : SHPHandle;
   cdecl;  external LibName name '_SHPOpen';

Function SHPCreate (pszShapeFile: Pchar; nShapeType: LongInt ) : SHPHandle;
   cdecl;  external LibName name '_SHPCreate';
Procedure SHPGetInfo (hSHP: SHPHandle; pnEntities: PLongInt; pnShapeType:
  PLongInt; padfMinBound: PDouble; padfMaxBound: PDouble ) ;  cdecl;  external
  LibName name '_SHPGetInfo';
Function SHPReadObject (hSHP: SHPHandle; iShape: LongInt ) : PShpObject;
   cdecl;  external LibName name '_SHPReadObject';
Function SHPWriteObject (hSHP: SHPHandle; iShape: LongInt; psObject: PShpObject
  ) : LongInt;  cdecl;  external LibName name '_SHPWriteObject';
Procedure SHPDestroyObject (psObject: PShpObject ) ;  cdecl;  external
  LibName name '_SHPDestroyObject';
Procedure SHPComputeExtents (psObject: PShpObject ) ;  cdecl;  external
  LibName name '_SHPComputeExtents';

//------------------------------------------------------------------------------
// SHP Create Object
//------------------------------------------------------------------------------
{** To Create a shape object of certain type

Function SHPCreateObject(    nSHPType: LongInt;
                             nShapeId: LongInt;
                             nParts:   LongInt;
                             panPartStart: PLongInt;
                             panPartType: PLongInt;
                             nVertices: LongInt;
                             padfX: PDouble;
                             padfY: PDouble;
                             padfZ: PDouble;
                             padfM: PDouble ) : PShpObject;

  nSHPType:		The SHPT_ type of the object to be created, such
                        as SHPT_POINT, or SHPT_POLYGON.

  iShape:		The shapeid to be recorded with this shape.
                        The number -1 means a new one.

  nParts:		The number of parts for this object.  If this is
                        zero for ARC, or POLYGON type objects, a single
                        zero valued part will be created internally.

  panPartStart:		The list of zero based start vertices for the rings
                        (parts) in this object.  The first should always be
                        zero.  This may be NULL if nParts is 0.

  panPartType:		The type of each of the parts.  This is only meaningful
                        for MULTIPATCH files.  For all other cases this may
                        be NULL, and will be assumed to be SHPP_RING.

  nVertices:		The number of vertices being passed in padfX,
                        padfY, and padfZ.

  padfX:		An array of nVertices X coordinates of the vertices
                        for this object.

  padfY:		An array of nVertices Y coordinates of the vertices
                        for this object.

  padfZ:		An array of nVertices Z coordinates of the vertices
                        for this object.  This may be NULL in which case
		        they are all assumed to be zero.

  padfM:		An array of nVertices M (measure values) of the
			vertices for this object.  This may be NULL in which
			case they are all assumed to be zero.
}
Function SHPCreateObject ( nSHPType: LongInt;
                           nShapeId: LongInt; nParts: LongInt;
                           panPartStart: PLongInt;
                           panPartType: PLongInt;
                           nVertices: LongInt;
                           padfX: PDouble;
                           padfY: PDouble;
                           padfZ: PDouble;
                           padfM: PDouble ) : PShpObject;  cdecl;  external LibName name
    '_SHPCreateObject';

//------------------------------------------------------------------------------
// SHP Create Siple Object
//------------------------------------------------------------------------------
{**

  Function SHPCreateSimpleObject (  nSHPType:  LongInt;
                                    nVertices: LongInt;
                                    padfX:  PDouble;
                                    padfY: PDouble;
                                    padfZ: PDouble ) : PShpObject;

  nSHPType:		The SHPT_ type of the object to be created, such
                        as SHPT_POINT, or SHPT_POLYGON.

  nVertices:		The number of vertices being passed in padfX,
                        padfY, and padfZ.

  padfX:		An array of nVertices X coordinates of the vertices
                        for this object.

  padfY:		An array of nVertices Y coordinates of the vertices
                        for this object.

  padfZ:		An array of nVertices Z coordinates of the vertices
                        for this object.  This may be NULL in which case
		        they are all assumed to be zero.

}
Function SHPCreateSimpleObject (nSHPType: LongInt; nVertices: LongInt; padfX:
  PDouble; padfY: PDouble; padfZ: PDouble ) : PShpObject;  cdecl;  external
  LibName name '_SHPCreateSimpleObject';

//------------------------------------------------------------------------------
// Close a Shapefile by a given handle
//------------------------------------------------------------------------------
Procedure SHPClose (hSHP: SHPHandle ) ;  cdecl;  external LibName name
  '_SHPClose';

//------------------------------------------------------------------------------
// Get back the shape type sting
//------------------------------------------------------------------------------
Function SHPTypeName (nSHPType: LongInt ) : Pchar;  cdecl;  external
  LibName name '_SHPTypeName';

//------------------------------------------------------------------------------
// Get back the part type sting
//------------------------------------------------------------------------------
Function SHPPartTypeName (nPartType: LongInt ) : Pchar;  cdecl;  external
  LibName name '_SHPPartTypeName';

// ----------------------------------------------------------------------------
// Shape quadtree indexing API.
// .. this can be two or four for binary or quad tree
// ----------------------------------------------------------------------------

Const
  MAX_SUBNODE = 4;

// ----------------------------------------------------------------------------
// region covered by this node list of shapes stored at this node.  
// The papsShapeObj pointers or the whole list can be NULL  
// ----------------------------------------------------------------------------
Type
  PSHPTreeNode = ^SHPTreeNode;
  SHPTreeNode = Record
    adfBoundsMin: Array[0..3] Of Double;
    adfBoundsMax: Array[0..3] Of Double;
    nShapeCount: LongInt;
    panShapeIds: PLongInt;
    papsShapeObj: PPShpObject;
    nSubNodes: LongInt;
    apsSubNode: Array[0.. (MAX_SUBNODE ) - 1] Of PSHPTreeNode;
  End;

  shape_tree_node = SHPTreeNode;
  Pshape_tree_node = ^shape_tree_node;

  PSHPTree = ^SHpTree;
  SHpTree = Record
    hSHP: SHPHandle;
    nMaxDepth: LongInt;
    nDimension: LongInt;
    psRoot: PSHPTreeNode;
  End;

Function SHPCreateTree (hSHP: SHPHandle; nDimension: LongInt; nMaxDepth:
  LongInt; padfBoundsMin: PDouble; padfBoundsMax: PDouble ) : PSHPTree; 
  cdecl;  external LibName name '_SHPCreateTree';
Procedure SHPDestroyTree (hTree: PSHPTree ) ;  cdecl;  external LibName
  name '_SHPDestroyTree';
Function SHPWriteTree (hTree: PSHPTree; pszFilename: Pchar ) : LongInt; 
  cdecl;  external LibName name '_SHPWriteTree';
Function SHPReadTree (pszFilename: Pchar ) : SHpTree;  cdecl;  external
  LibName name '_SHPReadTree';
Function SHPTreeAddObject (hTree: PSHPTree; psObject: PShpObject ) : LongInt;
   cdecl;  external LibName name '_SHPTreeAddObject';
Function SHPTreeAddShapeId (hTree: PSHPTree; psObject: PShpObject ) : LongInt;
   cdecl;  external LibName name '_SHPTreeAddShapeId';
Function SHPTreeRemoveShapeId (hTree: PSHPTree; nShapeId: LongInt ) : LongInt;
   cdecl;  external LibName name '_SHPTreeRemoveShapeId';
Procedure SHPTreeTrimExtraNodes (hTree: PSHPTree ) ;  cdecl;  external
  LibName name '_SHPTreeTrimExtraNodes';
Function SHPTreeFindLikelyShapes (hTree: PSHPTree; padfBoundsMin: PDouble;
  padfBoundsMax: PDouble; _para4: PLongInt ) : PLongInt;  cdecl;  external
  LibName name '_SHPTreeFindLikelyShapes';
Function SHPCheckBoundsOverlap (_para1: PDouble; _para2: PDouble; _para3:
  PDouble; _para4: PDouble; _para5: LongInt ) : LongInt;  cdecl;  external
  LibName name '_SHPCheckBoundsOverlap';

// ----------------------------------------------------------------------------
// DBF Support.                              }
// ----------------------------------------------------------------------------

Type
  PDBFInfo = ^DbfInfo;
  DbfInfo = Record
    fp: File;
    nRecords: LongInt;
    nRecordLength: LongInt;
    nHeaderLength: LongInt;
    nFields: LongInt;
    panFieldOffset: PLongInt;
    panFieldSize: PLongInt;
    panFieldDecimals: PLongInt;
    pachFieldType: Pchar;
    pszHeader: Pchar;
    nCurrentRecord: LongInt;
    bCurrentRecordModified: LongInt;
    pszCurrentRecord: Pchar;
    bNoHeader: LongInt;
    bUpdated: LongInt;
  End;

  DBFHandle = PDBFInfo;

  DBFFieldType = (FTString, FTInteger, FTDouble, FTInvalid ) ;

Const  XBASE_FLDHDR_SZ = 32;

Function DBFOpen (pszDBFFile: Pchar; pszAccess: Pchar ) : DBFHandle;
  cdecl;  external LibName name '_DBFOpen';
Function DBFCreate (pszDBFFile: Pchar ) : DBFHandle;
  cdecl;  external LibName name '_DBFCreate';
Function DBFGetFieldCount (psDBF: DBFHandle ) : LongInt;
  cdecl;  external LibName name '_DBFGetFieldCount';
Function DBFGetRecordCount (psDBF: DBFHandle ) : LongInt;
  cdecl;  external LibName name '_DBFGetRecordCount';
Function DBFAddField (hDBF: DBFHandle; pszFieldName: Pchar; eType:
  DBFFieldType; nWidth: LongInt; nDecimals: LongInt ) : LongInt;
  cdecl;  external LibName name '_DBFAddField';
Function DBFGetFieldInfo (psDBF: DBFHandle; iField: LongInt; pszFieldName:
  Pchar; pnWidth: PLongInt; pnDecimals: PLongInt ) : DBFFieldType;
  cdecl;  external LibName name '_DBFGetFieldInfo';
Function DBFGetFieldIndex (psDBF: DBFHandle; pszFieldName: Pchar ) : LongInt;
  cdecl;  external LibName name '_DBFGetFieldIndex';
Function DBFReadIntegerAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt ) : LongInt;  cdecl;  external LibName name
  '_DBFReadIntegerAttribute';
Function DBFReadDoubleAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt ) : Double;  cdecl;  external LibName name
  '_DBFReadDoubleAttribute';
Function DBFReadStringAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt ) : Pchar;  cdecl;  external LibName name
  '_DBFReadStringAttribute';
Function DBFIsAttributeNULL (hDBF: DBFHandle; iShape: LongInt; iField: LongInt
  ) : LongInt;  cdecl;  external LibName name '_DBFIsAttributeNULL';
Function DBFWriteIntegerAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt; nFieldValue: LongInt ) : LongInt;  cdecl;  external LibName
  name '_DBFWriteIntegerAttribute';
Function DBFWriteDoubleAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt; dFieldValue: Double ) : LongInt;  cdecl;  external LibName name
  '_DBFWriteDoubleAttribute';
Function DBFWriteStringAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt; pszFieldValue: Pchar ) : LongInt;  cdecl;  external LibName
  name '_DBFWriteStringAttribute';
Function DBFWriteNULLAttribute (hDBF: DBFHandle; iShape: LongInt; iField:
  LongInt ) : LongInt;  cdecl;  external LibName name
  '_DBFWriteNULLAttribute';
Function DBFReadTuple (psDBF: DBFHandle; hEntity: LongInt ) : Pchar; 
  cdecl;  external LibName name '_DBFReadTuple';
Function DBFWriteTuple (psDBF: DBFHandle; hEntity: LongInt; Var pRawTuple ) :
  LongInt;  cdecl;  external LibName name '_DBFWriteTuple';
Function DBFCloneEmpty (psDBF: DBFHandle; pszFilename: Pchar ) : DBFHandle;
   cdecl;  external LibName name '_DBFCloneEmpty';
Procedure DBFClose (hDBF: DBFHandle ) ;  cdecl;  external LibName name
  '_DBFClose';
Function DBFGetNativeFieldType (hDBF: DBFHandle; iField: LongInt ) : char;
   cdecl;  external LibName name '_DBFGetNativeFieldType';

Implementation

End.
// ----------------------------------------------------------------------------
// EndOf 
// ----------------------------------------------------------------------------

