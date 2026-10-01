import { Type } from "class-transformer";
import {
  ArrayMaxSize,
  ArrayMinSize,
  IsArray,
  IsEnum,
  IsNotEmpty,
  IsOptional,
  IsString,
  Matches,
  MaxLength,
  ValidateIf,
  Validate,
  ValidateNested,
  ValidatorConstraint,
  ValidatorConstraintInterface,
  ValidationArguments,
} from "class-validator";
import { BATCH_LIMITS } from "./batch.constants";

export enum BatchEntryType {
  FILE = "file",
  TEXT = "text",
}

@ValidatorConstraint({ name: "entryShape", async: false })
class EntryShapeConstraint implements ValidatorConstraintInterface {
  validate(_value: BatchEntryType, args: ValidationArguments): boolean {
    const entry = args.object as BatchEntryDto;
    return entry.type === BatchEntryType.FILE
      ? typeof entry.fileRef === "string" && entry.text === undefined
      : entry.type === BatchEntryType.TEXT &&
          typeof entry.text === "string" &&
          entry.fileRef === undefined;
  }

  defaultMessage(): string {
    return "file entries require only fileRef and text entries require only text";
  }
}

@ValidatorConstraint({ name: "batchEntries", async: false })
class BatchEntriesConstraint implements ValidatorConstraintInterface {
  validate(value: BatchEntryDto[]): boolean {
    if (!Array.isArray(value)) return false;
    const ids = value.map((entry) => entry.id);
    const fileRefs = value
      .filter((entry) => entry.type === BatchEntryType.FILE)
      .map((entry) => entry.fileRef);

    return (
      new Set(ids).size === ids.length &&
      new Set(fileRefs).size === fileRefs.length &&
      fileRefs.length <= BATCH_LIMITS.maxFiles
    );
  }

  defaultMessage(): string {
    return `entry ids and fileRefs must be unique, with at most ${BATCH_LIMITS.maxFiles} files`;
  }
}

export class BatchEntryDto {
  @IsString()
  @IsNotEmpty()
  @MaxLength(80)
  @Matches(/^[a-zA-Z0-9_-]+$/)
  id!: string;

  @IsEnum(BatchEntryType)
  @Validate(EntryShapeConstraint)
  type!: BatchEntryType;

  @IsString()
  @IsNotEmpty()
  @MaxLength(120)
  store!: string;

  @ValidateIf((entry: BatchEntryDto) => entry.type === BatchEntryType.FILE)
  @IsString()
  @IsNotEmpty()
  @MaxLength(120)
  @Matches(/^[a-zA-Z0-9_-]+$/)
  fileRef?: string;

  @ValidateIf((entry: BatchEntryDto) => entry.type === BatchEntryType.TEXT)
  @IsString()
  @IsNotEmpty()
  @MaxLength(BATCH_LIMITS.maxTextCharacters)
  text?: string;
}

export class ProcessBatchDto {
  @IsOptional()
  @IsString()
  @MaxLength(120)
  name?: string;

  @IsArray()
  @ArrayMinSize(1)
  @ArrayMaxSize(BATCH_LIMITS.maxEntries)
  @Validate(BatchEntriesConstraint)
  @ValidateNested({ each: true })
  @Type(() => BatchEntryDto)
  entries!: BatchEntryDto[];
}
