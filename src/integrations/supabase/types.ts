export type Json = string | number | boolean | null | { [key: string]: Json | undefined } | Json[];

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.5";
  };
  public: {
    Tables: {
      generated_reports: {
        Row: {
          created_at: string;
          dates: Json;
          file_hash: string;
          file_name: string;
          file_size: number;
          id: string;
          local_history_id: string | null;
          matched_employees: number;
          pdf_count: number;
          report_key: string;
          report_type: string;
          total_employees: number;
          updated_at: string;
        };
        Insert: {
          created_at?: string;
          dates?: Json;
          file_hash: string;
          file_name: string;
          file_size?: number;
          id?: string;
          local_history_id?: string | null;
          matched_employees?: number;
          pdf_count?: number;
          report_key: string;
          report_type?: string;
          total_employees?: number;
          updated_at?: string;
        };
        Update: {
          created_at?: string;
          dates?: Json;
          file_hash?: string;
          file_name?: string;
          file_size?: number;
          id?: string;
          local_history_id?: string | null;
          matched_employees?: number;
          pdf_count?: number;
          report_key?: string;
          report_type?: string;
          total_employees?: number;
          updated_at?: string;
        };
        Relationships: [];
      };
      employees: {
        Row: {
          created_at: string;
          employee_code: string | null;
          id: string;
          name: string;
          status: string;
          updated_at: string;
        };
        Insert: {
          created_at?: string;
          employee_code?: string | null;
          id?: string;
          name: string;
          status?: string;
          updated_at?: string;
        };
        Update: {
          created_at?: string;
          employee_code?: string | null;
          id?: string;
          name?: string;
          status?: string;
          updated_at?: string;
        };
        Relationships: [];
      };
      error_logs: {
        Row: {
          created_at: string;
          error_message: string;
          file_name: string | null;
          id: string;
        };
        Insert: {
          created_at?: string;
          error_message: string;
          file_name?: string | null;
          id?: string;
        };
        Update: {
          created_at?: string;
          error_message?: string;
          file_name?: string | null;
          id?: string;
        };
        Relationships: [];
      };
      report_files: {
        Row: {
          file_hash: string;
          file_name: string;
          id: string;
          processed_status: string;
          report_date: string | null;
          upload_date: string;
        };
        Insert: {
          file_hash: string;
          file_name: string;
          id?: string;
          processed_status?: string;
          report_date?: string | null;
          upload_date?: string;
        };
        Update: {
          file_hash?: string;
          file_name?: string;
          id?: string;
          processed_status?: string;
          report_date?: string | null;
          upload_date?: string;
        };
        Relationships: [];
      };
      reports: {
        Row: {
          created_at: string;
          date: string;
          file_name: string;
          id: string;
        };
        Insert: {
          created_at?: string;
          date: string;
          file_name: string;
          id?: string;
        };
        Update: {
          created_at?: string;
          date?: string;
          file_name?: string;
          id?: string;
        };
        Relationships: [];
      };
      selfie_records: {
        Row: {
          count: number;
          created_at: string;
          date: string;
          employee_code: string | null;
          employee_name: string;
          id: string;
          source_file_id: string | null;
        };
        Insert: {
          count?: number;
          created_at?: string;
          date: string;
          employee_code?: string | null;
          employee_name: string;
          id?: string;
          source_file_id?: string | null;
        };
        Update: {
          count?: number;
          created_at?: string;
          date?: string;
          employee_code?: string | null;
          employee_name?: string;
          id?: string;
          source_file_id?: string | null;
        };
        Relationships: [];
      };
    };
    Views: {
      [_ in never]: never;
    };
    Functions: {
      [_ in never]: never;
    };
    Enums: {
      [_ in never]: never;
    };
    CompositeTypes: {
      [_ in never]: never;
    };
  };
};

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">;

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">];

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R;
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] & DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R;
      }
      ? R
      : never
    : never;

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I;
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I;
      }
      ? I
      : never
    : never;

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U;
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U;
      }
      ? U
      : never
    : never;

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never;

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never;

export const Constants = {
  public: {
    Enums: {},
  },
} as const;
