﻿using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace DBUtilityHub.Migrations
{
    /// <inheritdoc />
    public partial class refer : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_ColumnMetaDataEntity_ColumnMetaDataEntity_ReferenceColumnMe~",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropForeignKey(
                name: "FK_ColumnMetaDataEntity_TableMetaDataEntity_ReferenceTableMeta~",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropForeignKey(
                name: "FK_ColumnMetaDataEntity_TableMetaDataEntity_TableMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropIndex(
                name: "IX_ColumnMetaDataEntity_ReferenceColumnMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropIndex(
                name: "IX_ColumnMetaDataEntity_ReferenceTableMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropIndex(
                name: "IX_ColumnMetaDataEntity_TableMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropColumn(
                name: "ReferenceColumnMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropColumn(
                name: "ReferenceTableMetaDataId",
                table: "ColumnMetaDataEntity");

            migrationBuilder.DropColumn(
                name: "TableMetaDataId",
                table: "ColumnMetaDataEntity");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "ReferenceColumnMetaDataId",
                table: "ColumnMetaDataEntity",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "ReferenceTableMetaDataId",
                table: "ColumnMetaDataEntity",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "TableMetaDataId",
                table: "ColumnMetaDataEntity",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateIndex(
                name: "IX_ColumnMetaDataEntity_ReferenceColumnMetaDataId",
                table: "ColumnMetaDataEntity",
                column: "ReferenceColumnMetaDataId");

            migrationBuilder.CreateIndex(
                name: "IX_ColumnMetaDataEntity_ReferenceTableMetaDataId",
                table: "ColumnMetaDataEntity",
                column: "ReferenceTableMetaDataId");

            migrationBuilder.CreateIndex(
                name: "IX_ColumnMetaDataEntity_TableMetaDataId",
                table: "ColumnMetaDataEntity",
                column: "TableMetaDataId");

            migrationBuilder.AddForeignKey(
                name: "FK_ColumnMetaDataEntity_ColumnMetaDataEntity_ReferenceColumnMe~",
                table: "ColumnMetaDataEntity",
                column: "ReferenceColumnMetaDataId",
                principalTable: "ColumnMetaDataEntity",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_ColumnMetaDataEntity_TableMetaDataEntity_ReferenceTableMeta~",
                table: "ColumnMetaDataEntity",
                column: "ReferenceTableMetaDataId",
                principalTable: "TableMetaDataEntity",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_ColumnMetaDataEntity_TableMetaDataEntity_TableMetaDataId",
                table: "ColumnMetaDataEntity",
                column: "TableMetaDataId",
                principalTable: "TableMetaDataEntity",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
